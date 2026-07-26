# Code Review Findings — July 2026

Full output of a five-part review of SharePointSmartCopy, run 2026-07-25/26 against commit `3171d83`.
This is a **work queue**, not a narrative: each finding is self-contained and independently
implementable. Nothing here is fixed unless its status says so.

## Status as of 2026-07-26

28 findings fixed (uncommitted, on top of `3171d83`), verified by a clean build after every change.
**None of this has been run against a live SharePoint tenant** — everything below was implemented and
build-verified only; test before shipping.

**✅ Fixed:** A1, A3, A4, A5, A6, A8, A9 (primary — see note), A10, A11, A12, A13, B1, B2, B3 (Enhanced
REST side only), B4 (partial — see note), B5, B8, B9, B10, B12, B13, B14 (documented, not hardened —
see note), C1, C3, C4, C5 (gate half only), C8, C9.

**Deferred — not attempted.** A2, C2, C6, and C5's re-walk removal all hinge on one restructuring
(making the scan emit folder identities and feeding that to both engines' folder-metadata passes,
replacing SPMI's ancestor-hop scheme). That's a rewrite of the single most heavily-used code path in
the app, with an unresolved API constraint (`$expand=listItem` unsupported on `/delta`, see C2) and no
live tenant to verify against — attempting it blind risked silently corrupting folder metadata across
production tenants. B6, B7, B11, C7 (full incremental version), and C10 were also left open — each
needed closer reasoning than was safe to give in the same pass as everything else. Section D (special
containers) is unstarted new-feature work, out of scope for this pass.

## How to read this

- **CONFIRMED** — traced in the code; it definitely misbehaves as described.
- **PLAUSIBLE** — depends on SharePoint/Graph behavior that could not be verified without a live
  tenant. Verify before investing in a fix.
- **Effort** — rough implementation size. **Risk** — chance of breaking something else.

Findings are grouped by theme and ordered by severity within each group. IDs are stable; reference
them in commits.

**Two standing cautions.** First, this codebase's comments document a long history of real incidents
and their fixes; before changing anything, read the comment above it — several findings below exist
precisely because a documented invariant was later violated. Second, several findings note that a
failure is currently *silent*; when fixing, preserve the distinction the code already draws between
"nothing to do", "failed", and "partially succeeded" rather than collapsing them.

---

## A. Silent data loss and status misreporting

The highest-priority group. In every case the app tells the user something untrue — a Success that
isn't, or a silent omission — which is worse than a visible failure because it defeats reconciliation.

### A1 — Permission-flag bulk read truncates; items silently copy with inherited permissions
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`SharePointService.BulkReadPermissionFlagsAsync` (~`:4885` `if (!resp.IsSuccessStatusCode) break;`
and ~`:4897` `catch { break; }`) returns the **partial** dictionary when a page fails mid-pagination.
`CopyService.cs:419` then does `if (!permissionFlags.TryGetValue(flagKey, out var hu) || !hu) return;`
— a missing key is indistinguishable from "this item has no unique permissions".

Failure scenario: a library with >5,000 items (`$select=HasUniqueRoleAssignments` is an unindexed
computed property, so exactly the query SharePoint throttles hardest); page 2 returns 429 after the
retries are exhausted → every item past ~5,000 copies with inherited permissions, no row logged, run
reports Success.

Fix: throw instead of `break`, matching the fix already applied to the sibling methods
`GetListItemsAsync` (~`:4628`) and `GetListItemTitlesAsync` (~`:4677`) — whose comment reads *"Never
truncate silently: a failed page used to `break`…"*. The fix was simply never applied here.

### A2 — SPMI stamps folders *outside the copy scope*, including the target library root
**CONFIRMED · Effort: M · Risk: Medium**

`MigrationJobService.cs:514-543` computes `hopsUp = bestDepth - pathDepth` from **target** path depth,
then `SharePointService.FetchFolderMetadataAsync` (~`:1348-1359`) walks that many **source**
`parentReference` hops. Any target-side-only path segment shifts the walk by that many levels.

Failure scenario: target folder `Docs/Archive` selected (so `TargetSubFolderPath="Archive"`), source
folder `Q1` copied. Ancestor key `"Archive"` has no direct files → borrows the shallowest descendant
`"Archive/Q1"` → `hopsUp=1` → walks up from source `Q1` to the **source library root** → then
`PatchFolderMetadataAsync` overwrites the pre-existing target `Docs/Archive` folder's
Created/Modified/Author/Editor with it. Separately, key `""` (source library root metadata) always
patches the **target library root** itself. Neither folder is part of the copy, the change is
irreversible, and it happens on every SPMI run including the `groupTasks.Count == 0` "cheap repair"
path (~`:700-710`).

Enhanced REST's equivalent pass is correct here — it only ever stamps `prefix + SourceName` and
`prefix + relativePath`, never the prefix alone. The two engines disagree; make SPMI match.

### A3 — A blank page is reported Success
**CONFIRMED (verified directly) · Effort: S · Risk: Low · ✅ FIXED**

`CopyService.cs:755-758`: when `pageMeta == null`, the code sets `result.ErrorMessage` and falls
through; `CopySingleFileAsync` then unconditionally sets `result.Status = CopyStatus.Success` at
`:611`. The documented earlier fix covered only the `saveErr != null` branch (`:742` throws).

Failure scenario: `GetPageMetadataAsync` fails (throttled REST call, or `job.SourceSiteUrl` empty at
`:715`) → `CreatePageStubAsync` has already created an empty `.aspx` at the target → row shows
Success, counts in `FileSuccessCount`, lands in the Excel report as a success, and a Copy-If-Newer
re-run skips it because the stub is newer than the source.

Fix: throw in that branch too, matching `:742`.

### A4 — A fully-copied file is reported Cancelled
**CONFIRMED (verified directly) · Effort: S · Risk: Low · ✅ FIXED**

`CopyService.cs:643` — the per-file permission block's `catch (OperationCanceledException) { throw; }`
escapes to the outer handler at `:647`, which sets `Status = Cancelled, ErrorMessage = "Cancelled"`,
overwriting the `Success` set at `:611` (or the `Skipped`/"Up to date" from `:590`).

The file is fully present at the target, but the grid and the saved report classify it as "never
actually attempted" (per `CopyResult.cs:8-11`), and `CancelledCount` is inflated.

Fix: capture the status before the permission block and don't let a post-copy cancellation downgrade
a completed copy; report the cancellation separately.

### A5 — Lookup-value cache poisons a value for the whole session
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`SharePointService.ResolveLookupValueAsync` (~`:3805-3819`): `result` stays `null` on a non-success
response (post-retry 429/500/401) and is then written to `_lookupValueCache` **unconditionally**
(~`:3818`). The cache is never cleared for the app's lifetime. `ApplyFileCustomFieldsAsync` (~`:3881`)
then does `if (resolvedIds.Count == 0) continue;` — the field is skipped and **not** added to
`lookupErrors`, so the file reports Success with a blank Lookup column.

Failure scenario: one throttled probe for display value "Acme Corp" poisons it for the rest of the
session; all 40,000 files referencing it lose that column silently. Partial resolution is equally
silent — 3 of 4 multi-lookup entries resolving writes only 3, with no warning.

Fix: cache only successful lookups (negative caching is fine for a genuine "no match", but not for a
transport failure — distinguish them), and add unresolved values to `lookupErrors`.

### A6 — Folder-metadata pass reports "complete" after abandoning every remaining folder
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`CopyService.ApplyAllFolderMetadataAsync:514-527`: `catch { }` at `:525` leaves `completed = true`,
then `onDone?.Report(completed)`. The `foreach` at `:519` is sequential and the inner
`Parallel.ForEachAsync` (`:1018`) cancels its siblings on first exception.

Failure scenario: `GetOrCreateFolderPathAsync` or `PatchFileSystemDateAsync` throws on folder #1 of
3,000 (throttle retries exhausted) → 2,999 folders never stamped → wizard reports "Folder metadata
updated in 0m 4s" and lets the user proceed to the report. The method takes no `activityLog`, so
there is no log line either. Only `OperationCanceledException` is reported as incomplete.

Fix: set `completed = false` in the general catch, pass an `activityLog` so the failure is visible,
and consider per-folder isolation so one folder's failure doesn't abandon the rest.

### A7 — `ValidateUpdateListItem` response parse failure is reported as success
**PLAUSIBLE · Effort: S · Risk: Low · PARTIALLY FIXED**

Any deviation in the response shape (`ErrorCode` as a string, `HasException` as `"false"`, a missing
`value` array) throws inside the `try`, is swallowed by `catch { }`, and the method returns `null` —
i.e. *"metadata applied successfully"*. HTTP was 200, so nothing else catches it.

**Status:** fixed in `PatchTimestampsViaRestAsync`. The same shape remains at ~`:3922-3937`
(`ApplyFileCustomFieldsAsync`) and ~`:4801-4814` (`ValidateUpdateItemFieldsAsync`).

Important related fact established during review: per `[MS-CSOMSPT]`, ValidateUpdateListItem's commit
is **all-or-nothing** — if any field raises an exception the item is not committed at all — and
`HasException: false` on a field means only that its value *parsed*, **not** that it was applied. Any
code treating a 200 or a clean per-field entry as proof of persistence is unsound.

### A8 — `GetRoleAssignmentsAsync` swallow-to-empty reports "permissions copied" when nothing was read
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

~`:4946` `return [];` and ~`:4968` `catch { return []; }`. `PermissionCopyService.cs:46` reads source
assignments and `:59-60` returns `new PermissionCopyResult(name, 0, [])` — success, zero applied, no
error — when the list is empty. A transient failure reading the source `roleassignments` is
indistinguishable from "the item has no non-LimitedAccess assignments".

Not destructive (the guard at `PermissionCopyService.cs:42` and the empty-check happen *before*
`BreakPermissionInheritance`), just silent. Fix: distinguish "read failed" from "genuinely empty".

### A9 — `GetCurrentVersionIdAsync` swallow-to-null leaves phantom versions, and drives a *delete*
**CONFIRMED (primary) / PLAUSIBLE (secondary) · Effort: S · Risk: Low · ✅ PRIMARY FIXED**

Primary fix: the swallow is now surfaced as a `result.ErrorMessage` instead of silently skipping the
delete. **Secondary (the newest-first ordering assumption) was NOT touched** — still open, still
unconfirmed without a live tenant.

`SharePointService.cs:2265-2273` — `catch { return null; }`. `CopyService.cs:871` captures the ID and
`:895-899` does `if (uploadVersionId != null)`, so a transient failure means the temporary "upload"
version is never deleted and **no error is recorded** (the `??=` at `:899` is unreachable). Every
replayed version of an affected file leaves a duplicate phantom entry; the file reports Success with
double the expected version count.

Secondary, **PLAUSIBLE and more serious**: `page?.Value?.FirstOrDefault()?.Id` assumes Graph returns
versions newest-first, with no `$orderby` and without the file's own `SortVersions` (~`:1249`, whose
comment documents that ordering as an assumption). If the order ever differs,
`DeleteItemVersionAsync` deletes an already-replayed *historical* version — real data loss. Worth
hardening regardless of whether it can be reproduced.

### A10 — `FetchFolderFileNamesRestAsync` returns partial pages, defeating the overwrite pre-flight
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`SharePointService.cs:2763-2771` — `if (!response.IsSuccessStatusCode) { Debug.WriteLine(...); return
result; }` conflates "folder 404 / genuinely empty" with "page 3 of 6 failed". The result is merged in
`MigrationJobService.cs:1165-1175` specifically to catch AllDocs rows Graph doesn't return.

Failure scenario: a folder with 12,000 existing files, page 2 returns 500 → those names are absent
from the snapshot → Overwrite mode skips their `PermanentlyDeleteFileAsync` calls → SPMI rejects each
with "already exists" — precisely the 2026-07-02 regression this method's own header says it was
written to prevent.

Fix: distinguish `StatusCode == NotFound` (return empty) from everything else (throw).

### A11 — `ChildExistsAsync` / `DeleteChildIfExistsAsync` swallow every failure as "not there"
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`SharePointService.cs:1489-1499` and `:1506-1518`, both `catch { return false; }`. `GetFileInfoAsync`
(~`:2859-2872`) documents exactly why this is wrong and does it properly — `catch (ODataError ex) when
(ex.ResponseStatusCode == 404)` only, with the comment *"swallowing every failure here made a
transient error … read as 'not there'"*. The folder equivalents never got the same treatment.
Additionally `CopyService.PrepareNativeCopyTargetAsync:151-159` **discards**
`DeleteChildIfExistsAsync`'s return value entirely (`await …; return true;`).

Failure scenario: a OneNote notebook or Document Set locked/checked-out in Overwrite mode — the delete
throws, is swallowed, the caller proceeds, and `CopyFolderNativeAsync` fails with `nameAlreadyExists`.
The user sees "already exists" instead of the real cause (locked / access denied).

### A12 — Site-scope page copies silently drop all custom column values (cache-key mismatch)
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`MainViewModel.cs:1987` passes `pageBulkCache` straight from `BulkReadCustomFieldsAsync`, which keys
by **bare item id** (`SharePointService.cs:3693-3697`), while `CopyService.cs:785` looks up
`$"{listId}:{listItemId}"`. Every lookup misses → no `ApplyFileCustomFieldsAsync` call, no error, row
reports Success. The two library paths (`:1671-1675`, `:1768-1769`) re-key correctly; this is specific
to Site-scope Site Pages.

### A13 — Version-cache pagination can truncate silently yet mark the item complete
**PLAUSIBLE · Effort: S · Risk: Low · ✅ FIXED**

`SharePointService.FetchBatchChunkAsync` ~`:1136-1144`: the `while (nextLink != null)` loop assigns
`nextLink = page?.OdataNextLink`, so a `null` page (empty/204 body from Kiota) exits the loop cleanly
with a **partial** version list, and the item is still added to `result` (~`:1151-1152`) as if
complete. A *thrown* page is handled correctly by `catch {}` at ~`:1148`; only the `null` page is
unguarded. The caller comment at ~`:934-936` states the consequence: *"an incomplete cache would
under-count a file's versions, and that file's batch could then exceed the SPMI entry ceiling and fail
import."*

### A14 — Cancelling mid-SPMI marks files "Cancelled" while SharePoint keeps importing them
**PLAUSIBLE (behavioral) · Effort: M · Risk: Low**

`MigrationJobService.cs:2097-2106` (also `:1091-1099`, `:1885-1893`). Cancel stops *our polling*; the
submitted import job continues server-side and lands the files. Rows say Cancelled, whose documented
meaning is "never actually attempted" (`CopyResult.cs:8-11`), so the saved report understates the
copy. Fix by reporting an explicit "submitted, outcome unknown" state, or by reconciling after cancel.

---

## B. Correctness and robustness

### B1 — "Re-apply folder metadata every run" is a no-op in Enhanced REST mode
**CONFIRMED (verified directly) · Effort: S · Risk: Low · ✅ FIXED**

`CopyService.cs:35` accepts `reapplyFolderMetadata`, but it is used **only** at `:391` (the migration
branch). The REST branch at `:475-500` never consults it: `:494` gates the folder pass on
`preserveMetadata && anyFileCopied`, and `applyMetadata:` is likewise `preserveMetadata &&
anyFileCopied`.

Failure scenario — exactly what the option's tooltip promises (`MainWindow.xaml:996`: *"needed to fix
folder metadata on an already-copied target"*): user re-runs Copy-If-Newer over a completed copy with
the box checked. Every file is up to date → `anyFileCopied` false → no folder pass runs at all, and
`onMetadataDone?.Report(true)` (`:500`) makes the wizard display "folder metadata complete" with 0
folders. Even when one file did copy, `dirtyFolderPaths` restricts the pass to that file's ancestor
chain, so the rest of the tree is still never repaired.

Fix: when `reapplyFolderMetadata` is set, run the pass regardless of `anyFileCopied` and pass
`dirtyFolderPaths: null`.

### B2 — Enhanced REST buffers whole files in memory, with no gate and a hard 2 GiB ceiling
**CONFIRMED · Effort: M · Risk: Medium · ✅ FIXED**

`CopyService.cs:763-768` and `:853-858`:
```csharp
using var stream = await spService.DownloadFileAsync(job.SourceDriveId, job.SourceItemId);
using var ms     = new MemoryStream();
await stream.CopyToAsync(ms, ct);
```
`maxParallel` × full file size resident, unbounded. SPMI has three layers of protection —
`largeFileGate` (2 slots), `TransferMemoryBudget` (~40% RAM), and the per-batch byte cap — all
constructed inside `MigrationJobService.ExecuteAsync` (~`:99-130`) and **unreachable from CopyService's
REST branch**. At 16 parallel copies of 1 GB files that's 16 GB live; the `qptiff` incident class is
fully reproducible here, with none of the fixes that incident produced.

Compounding correctness bug: `MigrationJobService.cs:1512-1517` fails >2 GB files with *"Copy this file
with Enhanced REST mode"* — but `MemoryStream` caps at `int.MaxValue`, so Enhanced REST throws
`IOException: Stream was too long` on the same file. **The documented escape hatch does not work.**

Fix, two parts: (a) hoist `TransferMemoryBudget` + a large-file semaphore into `CopyService.ExecuteAsync`
and share the instances with both engines (which also makes the budget genuinely process-global);
`job.SourceSize` is already carried from the scan (`:340`). (b) Spill above a threshold to a temp
`FileStream` with `FileOptions.DeleteOnClose` — still seekable, so `LargeFileUploadTask`
(~`:3129`) works unchanged, memory becomes O(1), and the 2 GiB ceiling disappears. New failure mode to
handle: available disk space.

Verified separately and **not** to be touched: the SPMI large-file gate can no longer be bypassed —
`MigrationJobService.cs:1509`/`:1536-1538` fall back to `job.SourceSize` and treat unknown sizes as
large. That fix is sound.

### B3 — Empty folders never receive their preserved dates/author, in either engine
**CONFIRMED · Effort: M · Risk: Low · ✅ PARTIALLY FIXED (Enhanced REST only)**

Enhanced REST: newly-created empty folders are now folded into `dirtyFolderPaths` so the metadata pass
picks them up. **SPMI side (`directFolderGroups`) is untouched** — tied to the deferred C6/A2 rewrite.

`CopyService.cs:284-330` creates them, but `dirtyFolderPaths` is built **only** from `allTasks` (i.e.
file jobs) at `:475-484`, and the subfolder filter at `:1005-1013` drops anything not in it. SPMI's
`directFolderGroups` (`MigrationJobService.cs:494-500`) is likewise keyed off file paths only.

Failure scenario: source tree with `Docs/Empty/` plus one changed file in `Docs/Other/`. `dirtyFolderPaths
= {"Docs","Docs/Other"}` → `Empty` is filtered out → created and reported Success, but keeps today's
date and the migrating account as Created By/Modified By. In SPMI it gets neither a manifest `<Folder>`
entry nor a REST correction. Same "empty folders are invisible to everything keyed off files" class as
the two previously-fixed empty-folder bugs.

### B4 — `CancellationToken` accepted but not threaded through
**CONFIRMED · Effort: S · Risk: Low · ✅ PARTIALLY FIXED**

`ApplyFileCustomFieldsAsync`, `CreateListItemAsync`, `ValidateUpdateItemFieldsAsync`,
`UpdateListItemAsync`, and `GetSharePointIdsAsync` (including its `Task.Delay`) now thread `ct`
through properly, with cancellation no longer swallowed by the retry loop. `GetOrCreateFolderAsync`/
`GetOrCreateFolderPathAsync` were **left as-is** — no `ct` param at all, called from many places; the
ripple to add one was judged lower-value than the others (smaller retry ceiling: 4 attempts × ≤12s,
not 8 × 120s).

All of these take a `ct` and then omit `cancellationToken: ct` on the actual request:
`ApplyFileCustomFieldsAsync` (~`:3905`), `CreateListItemAsync` (~`:4726`),
`ValidateUpdateItemFieldsAsync` (~`:4788`), `UpdateListItemAsync` (~`:4833`), plus
`GetSharePointIdsAsync` (~`:1624-1626`, `Task.Delay(attempt * 1500)` with no token) and
`GetOrCreateFolderAsync` (~`:3370`). `ApplyFileCustomFieldsAsync` even passes `ct` correctly to
`ResolveLookupValueAsync` (~`:3877`) — just not to its own POST.

`SendSharePointRequestAsync` retries up to 8 times with delays capped at 120 s (~`:3453`), so a single
uncancellable item can hold ~16 minutes of `Task.Delay` after the user presses Cancel.

### B5 — Cancellation not threaded into the folder-metadata enumeration
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`CopyService.cs:1000` → `SharePointService.EnumerateFoldersAsync` (~`:426-437`) takes **no**
`CancellationToken`, and the walk is serial and recursive. Cancel during the folder-metadata phase on a
3,000-folder library leaves `IsCancelable` true (button live) but nothing stops until the entire source
tree has been walked — the same "~30 silent minutes" cost the *file* scan fixed with `scanController`.
Note the dirty filter is applied *after* this enumeration (`:1005`), so the optimization never avoids
the expensive part. See also C5, which removes the re-walk entirely.

### B6 — Library/Site scope leaves the detached folder pass unobserved; a second run disposes its CTS
**CONFIRMED · Effort: M · Risk: Medium**

The Library/Site call sites (`MainViewModel.cs:1691`, `:1778`, `:1998`) pass **no** `onMetadataDone`,
so `IsUpdatingMetadata` stays false: the run is declared complete, `SaveReport()` is reachable, and the
app can be closed while folder dates/authors/permissions are still being written by the
fire-and-forget pass (`CopyService.cs:464`/`:495`). In a multi-library site copy each subsequent
`ExecuteAsync` calls `ResetFolderSegmentCache()` (`:111`) while the previous library's detached pass is
still resolving folder paths. And because navigation isn't gated, Back→Next starts a new run that calls
`_copyCts.Dispose()` (`MainViewModel.cs:1503`) — `Parallel.ForEachAsync` registers on that token, so
the running pass dies with `ObjectDisposedException`, which A6's `catch { }` turns into "completed".

`StartCopyAsync`'s `IsUpdatingMetadata` gate on `CanGoNext`/`Back` does prevent re-entry; the hole is
specific to the library/site path.

### B7 — Import-worker exception abandons the prep producer, leaking gate/budget reservations
**CONFIRMED · Effort: M · Risk: Medium**

`MigrationJobService.cs:1065-1077`: `await Task.WhenAll(importWorkers)` throwing skips `await
prepProducer` (`:1076`). The unguarded surface is `RetryBatchAfterConflictAsync` —
`GetFileUniqueIdAsync`/`PermanentlyDeleteFileAsync` inside its `Parallel.ForEachAsync` (~`:953-998`)
are not wrapped, unlike everything else in the file. If all workers die, nobody drains the bounded
`prepChannel`; the producer blocks forever inside `PrepareBatchAsync` holding `largeFileGate` slots and
`memoryBudget` charges. `ExecuteAsync` then returns and its `using` disposes
`downloadController`/`largeFileGate`/`uploadController`, so the detached producer's later
`Release()`/`WaitAsync` throw `ObjectDisposedException` on a task nobody observes. Same shape for the
inner consumer loop: a non-cancellation throw escaping `:1625-1790` leaves `producerTask` (`:1455`)
unawaited and blocked on `pipe.Writer.WriteAsync`.

### B8 — `AddPermissionRow`: blocking `Dispatcher.Invoke`, plus a name-only fallback that mis-attributes
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED (as a side effect of the C4 fix)**

C4's signature change (`CopyResult? row` passed directly by the caller) eliminated the name-only
fallback entirely, and the `Dispatcher.Invoke` call was replaced with direct property sets (which
self-marshal via `CopyResult.OnPropertyChanged`'s existing `BeginInvoke`).

`CopyService.cs:1130-1143`. (a) Synchronous `Dispatcher.Invoke` called from up to `maxParallel`
concurrent copy threads and from the 8-wide folder/SPMI permission passes — the exact "many background
threads blocking on the dispatcher queue" pattern the documented `UCEERR_RENDERTHREADFAILURE` fix
removed from `CopyResult.OnPropertyChanged` (`CopyResult.cs:29-47`). It is also redundant, since
setting those two properties already marshals itself. (b) The fallback
`results.FirstOrDefault(r => r.FileName == perm.ItemName)` defeats the "silently no-ops for folder
results" comment: a folder permission outcome whose `TargetPath` matches nothing (folders always pass
`$"{TargetSiteUrl}/{relativePath}"`, `:988`/`:1052`) gets stamped onto any *file* row with the same
name — and since `FileFailedCount` counts `PermissionStatus == Failed` (`MainViewModel.cs:975`), a
failed folder ACL makes an unrelated, successfully copied file read as failed.

### B9 — `GetPageMetadataAsync` builds an OData filter from a filename without URL-encoding
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`SharePointService.cs:4152` does `fileName.Replace("'", "''")` (correct OData quote escaping) but never
`Uri.EscapeDataString`s the literal before interpolating it into the query string at `:4157`
(`$filter=FileName eq '{escapedName}'`). A page named `Q3 #Review.aspx` truncates the request at the
`#` → HTTP 400 or no match; `&` splits the query; `+` decodes to a space. The method then returns
*"Page 'X' not found in source Site Pages library"* (`:4200`) — a wrong diagnosis. This is the one place
that interpolates a user-controlled *name* into a REST path without either `ServerRelativePathArg`
(~`:2883`) or `Uri.EscapeDataString`, so the file's repeated "entirely ID-based, so names with
`#`/`%`/`+` are unaffected" claim does not hold here. Note this also causes A3's blank page.

### B10 — Session-lifetime caches with no reset and no bound
**CONFIRMED · Effort: S · Risk: Low · ✅ PARTIALLY FIXED**

`_columnCache` (the correctness-sensitive one — stale Lookup `LookupListId` across runs) now has a
`ResetColumnCache()`, called at the start of every run alongside `ResetFolderSegmentCache()`.
`_spIdsCache`/`_lookupValueCache` growth was left unbounded — a memory-only concern, not correctness
(and `_lookupValueCache`'s poisoning behavior is separately fixed by A5).

Only `_folderSegmentTasks` has a reset (`ResetFolderSegmentCache`, correctly called from
`CopyService.cs:111` for both engines). `_spIdsCache` (~`:37`), `_columnCache` (~`:3525`), and
`_lookupValueCache` (~`:3528`) grow unbounded and persist across runs in the same app session.
`_spIdsCache` gains one entry per item on a 100k+ copy (~20-40 MB, from `TryCacheBatchSpIds` ~`:1201`).
More importantly `_columnCache` means that if the user adds or retypes a target column between two
runs without restarting, `ApplyFileCustomFieldsAsync` (~`:3841`) keeps resolving against the stale
definition — a Lookup whose `LookupListId` changed resolves to the old list. `GetLibraryColumnsAsync`
has a `skipCache` parameter that nothing in the repo ever passes as `true`.

### B11 — SPMI version ceiling has no per-file clamp
**PLAUSIBLE · Effort: M · Risk: Low**

`MigrationJobService.cs:819-838`: the `currentBatch.Count > 0` guard correctly lets an oversized *byte*
file form its own batch, but there is no clamp for **version count**. With "Copy all versions"
(`maxVersions == 0`, `CopyService.cs:382`), `VersionsOf` returns the raw count, so a 400-version
document becomes a one-file batch emitting ~400 `<File>` entries — above the 250 ceiling the comments
at `:712-736` describe as real and data-dependent. The batch fails,
`RetryBatchAfterConflictAsync` only fires for "already exists" aborts, and a re-run reproduces the
identical batch.

### B12 — Date reads not pinned to `InvariantCulture`
**PLAUSIBLE · Effort: S · Risk: Low · ✅ FIXED**

Date *writes* are correct everywhere (`ToUniversalTime()` + `InvariantCulture`). Date *reads* use bare
`DateTimeOffset.TryParse` at ~`:2085`, ~`:2087`, ~`:2783`, ~`:4697`. Since these are `DateTimeOffset`s
compared as instants (`DatesClose` ~`:2104`), local-vs-UTC is not a bug; the exposure is a
non-Gregorian default calendar (e.g. `ar-SA`) misparsing the ISO string → `null` → in
`FetchFolderFileNamesRestAsync` that flips an If-Newer decision. Worth pinning; no reproducible failure
identified.

### B13 — Two small UI/latent issues
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

Both: `StatusDisplay`/`StatusColor` now have a `Cancelled` arm, and the never-added-to-`results` gap
(now `resultsBySourcePath` in `CopyService.cs`) flushes to `pendingResults` defensively if a future
caller doesn't pre-seed a row.

`CopyResult.StatusDisplay` (`CopyResult.cs:70-78`) has no `Cancelled` arm, so cancelled rows render a
**blank** status cell (Pending gets "⏳ Pending"). And `CopyService.cs:168` —
`FindResult(...) ?? CreateResult(job)` builds a `CopyResult` that is never added to
`results`/`pendingResults`, so any non-folder job whose caller didn't pre-seed a row would copy
invisibly. All current call sites do pre-seed, so this is latent; it is also an O(n) scan (see C4).

### B14 — `FetchVersionCountsAsync` is unhardened (currently uncalled)
**CONFIRMED, latent · Effort: S · Risk: Low · ⚠️ DOCUMENTED, NOT HARDENED**

Left as dead code deliberately — it's unreachable today, so hardening it would be speculative work
with no behavioral payoff. Added an explicit warning comment on the method instead, so whoever wires
it up next hits the warning before the bug.

~`:768` lacks the retry rounds, adaptive gate, and `Throttled` subscription its two documented siblings
have, and defaults a failed sub-request to a version count of 1 — the exact under-count ~`:934-936`
warns causes SPMI import failure. It is **uncalled** anywhere in the repo. Either harden it or delete
it before someone wires it up.

---

## C. Performance

Ordered by expected real-world impact. Quantities assume a 100k-file / 5,000-folder library, the
scale this app is documented to target.

### C1 — Enhanced REST spends 2 Graph GETs per file just to decide "skip"
**CONFIRMED · Effort: M · Risk: Low-Medium · Biggest single win (~20×) · ✅ FIXED**

Implemented via a per-target-folder snapshot cache (`GetOrBuildFolderSnapshotAsync` in
`CopyService.cs`, `Lazy<Task<T>>`-deduped with no-cache-on-fault, reset per run) plus
`job.SourceModified` for the IfNewer comparison, falling back to a per-file Graph read only when it's
missing (individually-selected files, which never go through the scan).

`CopyService.cs:570` (`GetFileInfoAsync`) and `:579` (`GetFileMetadataAsync`). On an all-skip
Copy-If-Newer re-run these are the *entire* per-file cost: **200,000 round trips, ~100% waste**. With
the scan (~5,000 folder listings) the steady-state re-run is ~205,000 calls.

Two independent redundancies:
- **`:579` is pure duplication.** The scan already captured the source modified date into
  `job.SourceModified` (`:339`), which is consumed *only* by `MigrationJobService` (`:443`, `:1341`).
  `MigrationJobService.cs:410-418` documents this exact fix — *"decide skip-vs-copy with ZERO Graph
  calls"* — applied to SPMI and never back-ported.
- **`:570` is one GET per file where one listing per folder would do.**
  `SharePointService.FetchFolderItemsAsync` (~`:2694`) already returns `name → (ItemId, Modified)` for
  a whole folder, case-insensitively, retry-hardened, and is already used by the SPMI pre-flight
  (`MigrationJobService.cs:341-357`). 100k files across 5,000 folders → ~5,000 calls instead of
  100,000.

Fix: after the scan, build a target snapshot per distinct `TargetSubFolderPath` via
`FetchFolderItemsAsync` (bounded-parallel behind `CreateThrottleAwareGate`) and pass it into
`CopySingleFileAsync`; replace `:579` with `job.SourceModified ?? await GetFileMetadataAsync(...)`. The
snapshot's `ItemId` satisfies the `upToDateItemId` permission-refresh path unchanged. Care point: Skip
mode must keep its exists-check semantics — the snapshot gives that directly.

**Net: ~205k → ~10k round trips.** Cheaper and lower-risk than the parked `deltaLink` idea.

### C2 — The source scan is one Graph listing per folder; `/delta` replaces the lot
**CONFIRMED · Effort: L · Risk: Medium (~100× on the scan)**

`WalkFilesForCopyAsync` (~`:377-422`) → `GetChildrenWithMetadataAsync` (~`:610-634`): one listing per
folder, always, at a hardcoded concurrency of 8 (`CopyService.cs:136`), and every folder's `DriveItem`
payload is fetched then discarded except for files.

`GET /drives/{id}/root/delta?$select=id,name,file,folder,package,size,lastModifiedDateTime,parentReference`
returns the whole tree in large pages — **~20-50 calls for 100k items instead of 5,000** — with
`parentReference.path` for path reconstruction. It also emits *folder* items in the same stream, which
directly enables C4 and C5.

**Important constraint found during review:** `$expand=listItem` is **not supported on `/delta`**
(confirmed by a Microsoft maintainer; returns `400 "One of the provided arguments is not
acceptable."`). Since `IsSpecialContainer` needs `listItem.contentType` to detect Document Sets, a
delta-based scan must either fall back to the `package` facet alone (losing Document Set detection —
regression, see D1) or do a supplementary pass. **Resolve this before committing to C2.**

On the parked idea: the *persistence* half (storing `deltaLink` between runs) is the fiddly,
state-carrying part and is **not** where the win is. Full-traversal delta alone gets ~100× with no
persisted state, no staleness/invalidation semantics, and no "what if the target changed" questions. Do
that half first; treat `deltaLink` persistence as a separate optional follow-up.

### C3 — `GetFormDigestAsync` and `EnsureUserAsync` re-fetched per folder
**CONFIRMED · Effort: S · Risk: Low · Best effort-to-payoff ratio in this document · ✅ FIXED**

Digest cached per site with a 20-min TTL; `EnsureUserAsync` caches successes only (never a failure —
same reasoning as A5, a transient failure isn't "this account doesn't exist").

**Digest** (~`:2030`): not cached. Called once per `PatchFolderViaCsomAsync`, once per
`StampFolderColorAsync`, once per `PatchFolderProgIdAsync`. A form digest is site-scoped and valid
~30 minutes. On 5,000 folders needing correction: **5,000-15,000 unnecessary `_api/contextinfo` POSTs
where ~1 would do.**

**EnsureUser** (~`:5009`): not cached in `PatchFolderMetadataAsync` (~`:1768`, ~`:1784`). Up to 2 POSTs
per folder (already deduped when author == editor). Distinct author/editor emails across a library are
typically tens. On 5,000 folders: **up to 10,000 POSTs where ~30 would do.** Note
`PermissionCopyService` already caches this correctly per session (`PermissionCopyService.cs:12`,
`:140-144`) — the folder path just doesn't use that cache.

A folder needing correction currently pays ~6-8 calls (`GetFolderCurrentMetadata` + 2×EnsureUser +
digest + ProcessQuery + `VerifyFolderPersonFields`, plus color). Caching both drops it to ~4 — a >40%
cut on a phase documented as having taken **39 minutes** on a real run
(`MigrationJobService.cs:622-626`).

Fix: `ConcurrentDictionary<string, (string digest, DateTimeOffset fetchedUtc)>` keyed by site with a
~20 min TTL; `ConcurrentDictionary<string, int?>` keyed by `site|email`.

**Do not touch** the fast-path skip at ~`:1703-1717` — it is correct and effective (an already-correct
folder costs 1 call and exits).

### C4 — `AddPermissionRow` and `FindResult` are O(n²)
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`CopyService.cs:1130-1135` does an O(n) scan of the growing `ObservableCollection` per permission row —
and the call site at `:638` passes `result.TargetPath` so the method can go find `result`, **which is
in scope at the call site**. Same at `:437`. Worst case (all files have unique permissions) on 100k
files: ~5×10⁹ `OrdinalIgnoreCase` comparisons, plus the fallback name-only scan when the first misses.

Fix: change the signature to accept `CopyResult? row` and pass `result` directly at `:638`/`:437`; the
two folder call sites (`:988`, `:1052`) genuinely have no matching row — pass `null` and keep the
documented no-op (which also fixes B8b). `FindResult` (~`:1097`) has the same shape; fold it in with a
path→result dictionary built once.

### C5 — Enhanced REST folder pass re-walks the whole tree, serially, unthrottled
**CONFIRMED · Effort: M · Risk: Medium · ✅ GATE HALF FIXED**

The throttle-aware gate (problem 3) is now in place — `ApplyAllFolderMetadataAsync` creates a
`CreateThrottleAwareGate` and subscribes to `Throttled` for the duration, and the inner
`Parallel.ForEachAsync` acquires/releases it per folder. **The redundant re-walk itself (problems 1-2)
was NOT removed** — that requires C6's scan-emits-folders change, deferred.

`CopyService.cs:1000-1001` → `EnumerateFoldersAsync` (~`:426-437`) → `GetChildrenAsync` (~`:284-328`).
Three compounding problems:

1. **Redundant re-enumeration** — the scan already walked this tree.
2. **Serial** — `yield return` inside a recursive `await foreach` serializes to one Graph call in
   flight regardless of any gate; the exact pattern fixed in `EnumerateFilesForCopyAsync`, documented
   at ~`:354-358` as *"~30 minutes on a 3,000-folder library"*. This copy was never converted. It also
   calls `GetChildrenAsync`, which has **no `$select`** — full `DriveItem` payloads plus a
   `SharePointNode` + placeholder child allocated per item.
3. **No throttle protection at all** — `ApplyAllFolderMetadataAsync` is launched fire-and-forget at
   `:495`, i.e. after `ExecuteCoreAsync` returned and after the `finally` at `:102` already
   unsubscribed `onThrottled`. The inner `Parallel.ForEachAsync` (`:1018`) runs at raw `maxParallel`
   with no `AdaptiveParallelismController`, no `Throttled` subscription, and no
   `CreateThrottleAwareGate`. **Every other analysis phase in the codebase uses one.** It issues 4+
   calls per folder into a tenant the copy phase just finished depleting.

Fix: add the gate first — independently valuable and near-free. Then feed folder identities from the
scan (with C6) and delete the re-walk.

### C6 — SPMI folder metadata: 4-5 calls per folder, nearly all avoidable
**CONFIRMED · Effort: M · Risk: Medium**

`SharePointService.cs:1348-1359`. Per folder: 1 GET for the sample file's `parentReference`, + `hopsUp`
more GETs walking ancestors, + `GetFileMetadataAsync` (1), + `GetFolderProgIdAsync` →
`GetSharePointIdsAsync` (1; folder IDs are not in `_spIdsCache`) + SP REST ProgID read (1). **5,000
folders ≈ 20,000-25,000 round trips**, and `MigrationJobService.cs:502-543` builds an elaborate
shallowest-descendant search purely to feed the hop count.

Every input already existed in the scan — `WalkFilesForCopyAsync` visits each folder's `DriveItem`
(~`:406-415`) and discards it unless special or empty.

Fix:
- Emit a folder entry from the walk (the mechanism exists — `IsEmptyFolder`/`IsSpecialFolder` already
  do it) carrying `path → itemId`. The `parentReference` hop walk and the whole `hopsUp` machinery
  become dead code — **which also fixes A2**.
- Widen the scan's `$select` (~`:617`) to include `createdDateTime,createdBy,lastModifiedBy`. Same
  listing call, **no extra round trip** — and the per-folder `GetFileMetadataAsync` disappears too.
- `GetFolderProgIdAsync` (2 calls/folder) returns null for every plain folder, and by construction
  every folder reaching this method *is* plain (special containers are diverted to native copy in the
  scan loop, `CopyService.cs:238-279`, and never produce `TargetSubFolderPath` entries). Gate the probe
  on the scan's already-free `IsSpecialContainer` signal.

**Net: ~4-5 calls/folder → ~0.** Note `MigrationJobService.cs:502-543` becomes dead code — delete it
rather than optimizing it (it is currently O(folders²), ~25M `StartsWith` calls at 5,000 folders).

### C7 — Six LINQ counters rescanned on the UI thread every 400 ms
**CONFIRMED · Effort: S-M · Risk: Low**

`MainViewModel.cs:965-977` (counters), `:2293-2303` (`UpdateProgress`), `:1061-1063` (400 ms
`DispatcherTimer`). Each counter is `CopyResults.Count(predicate)` — a full scan. `UpdateProgress` does
one scan itself (`:2298`) then raises `OnPropertyChanged` for six more (`:2382-2392`), each re-evaluated
by the bound UI. At 250k rows that is **~1.75M predicate invocations per tick, ~4.4M/sec, on the UI
thread**, for the whole multi-hour run. Given the `UCEERR_RENDERTHREADFAILURE` history and the
`Dispatcher.BeginInvoke` backlog counter already in `CopyResult` (`:24-27`), starving the UI thread
this way is worth removing on its own merits.

Fix: maintain incremental counts — `CopyResult` already funnels every status change through its
`OnPropertyChanged` override. Cheaper interim: cache the tallies with a dirty flag, recompute at most
once per second. Care point: catch every transition, including the bulk `Status = Copying` loops at
`MigrationJobService.cs:86-87`.

### C8 — `GetCurrentVersionIdAsync` re-fetches a growing version list per replayed version
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`SharePointService.cs:2265-2273`: no `$top`, no `$select` — the full version collection, to read one
Id. Called *inside* the per-version replay loop, and each iteration has added ~2 versions (upload +
phantom), so payloads grow 1, 3, 5, … — **O(N²) bytes across the replay**. Also `GetVersionsAsync`
(~`:709-721`) fetches full `DriveItemVersion` objects with no `$select` when only `id`,
`lastModifiedDateTime`, `lastModifiedBy`, and `size` are read.

Fix: `Top = 1`, `Select = ["id"]` on the former; explicit `$select` on the latter. Small per file, but
free — and see A9, same method.

### C9 — `Task.WhenAll` over 250k simultaneously-created tasks
**CONFIRMED · Effort: S · Risk: Low · ✅ FIXED**

`CopyService.cs:447-451`: `allTasks.Select(...)` materializes one async state machine per file, all
launched, all immediately parked on `controller.WaitAsync` — 250k `Task` objects + state machines + a
250k-node `SemaphoreSlim` wait queue + a 250k-element array for `WhenAll`. Well over 100 MB of pure
scheduling overhead before any file transfers.

Fix: `Parallel.ForEachAsync` with `MaxDegreeOfParallelism = maxParallel`, the pattern used everywhere
else here. `AdaptiveParallelismController` still governs live width, so throttle behavior is unchanged.
Per-item exception handling is already fully contained inside `CopySingleFileAsync`, so fail-fast
semantics won't change behavior.

### C10 — Per-file path-string duplication
**CONFIRMED, second-order · Effort: S · Risk: Low**

`ComputeTargetSubFolder` (~`:1067-1079`) allocates a fresh `TargetSubFolderPath` per file even though
every file in a folder produces the identical string; `SourceDisplayPath`/`TargetDisplayPath` likewise
(`:342`, `:349`). With `allTasks` holding 250k `CopyJob`s plus 250k `CopyResult`s for the run's
duration, that's roughly 500 MB of jobs/results, a meaningful slice of it duplicate path strings. A memo
dictionary keyed by the directory portion collapses the subfolder strings to one instance per folder.

---

## D. Special containers and folder fidelity

Answers the question "what other folder-like objects does this app mishandle?". Two are already
handled correctly: **OneNote notebooks** (via the `package` facet) and **Document Sets** (via content
type prefix `0x0120D520`). Confirmed during review: the prefix check **does** catch derived types,
including **Video Sets** (asset-library videos, `0x0120D520A808`, a Document Set descendant) — so
asset-library video folders are already covered.

### D1 — Native `/copy` loses version history inside every special container
**CONFIRMED (documented) · Effort: M · Risk: Medium · Highest-impact item in this section**

Graph's `driveItem: copy` reference states verbatim: *"File versions are only retained when the
`includeAllVersionHistory` parameter is explicitly set to `true`. Otherwise, only the latest version is
copied."* `CopyFolderNativeAsync` (~`:1535`) sends only `parentReference` + `name`. **Every file inside
every OneNote notebook and Document Set is copied latest-version-only, silently, in both engines,
regardless of the user's version settings.**

Not a one-liner: the same reference documents *"A known issue occurs when the `includeAllVersionHistory`
request parameter is ignored if the `name` request parameter is also passed"* — and this code always
passes `name`.

Related documented constraints on the same call, worth surfacing to users: *"Metadata isn't retained
when a driveItem is copied, including system metadata and custom metadata"*; *"Permissions are not
retained… The copied driveItem inherits the permissions of the destination folder"*; cross-geo copy
unsupported with app-only auth; a 30,000-driveItem per-operation cap (the poll loop also caps at
~9.5 min). Cross-tenant `/copy` between two tenants cannot work with a single tenant-scoped token, so
special containers in a cross-tenant run fail outright.

### D2 — **Verify first:** does `/copy` actually preserve Document Set-ness / the package facet?
**UNRESOLVED — live-tenant test · Effort: S (test only)**

Graph's own statement that copy retains no system metadata **contradicts the assumption written into
`CopyFolderNativeAsync`**. Nobody has published a definitive result. The entire special-container
strategy rests on this. **Test this before investing in D1 or D3.**

If `/copy` does *not* preserve them, the native-copy path needs a post-copy repair (set content type
per D3, set ProgID per D6) — or the strategy needs rethinking.

### D3 — Folder content types are never preserved
**CONFIRMED · Effort: M · Risk: Low**

`GetOrCreateFolderCoreAsync` (~`:3389`) posts `new DriveItem { Name, Folder = new Folder() }` — no
content type, ever. Every rebuilt folder gets the library's default `Folder` (`0x0120`).
`PatchFolderMetadataAsync` writes only dates/authors/color. So any `0x0120`-derived custom container
loses its type binding and therefore its custom columns, views, and any formatting/flow keyed on
content type.

Detection is **free** — `listItem.contentType.id` is already in the walk's payload. Flag anything
starting with `0x0120` that is neither exactly `0x0120` nor already special. Two specific widenings
worth making: `0x0120D5` (**Document Collection Folder**, `_Hidden`, the documented *parent* of
Document Set) and `0x012002`/`0x012004` (Discussion / Summary Task — folder-derived, but list-only so
unreachable via the drive API).

Repair does **not** need `/copy`: create the folder as now, then set the content type on its list item.
`PATCH /sites/{s}/lists/{l}/items/{i}` with `{"contentType":{"id":…}}` — note the item itself, **not**
`/fields`; `ContentTypeId` as a *field* does not work. Community-attested, not documented (the listItem
update reference documents only field updates). **Trap:** PATCHing an ID not present in the target
library does **not** error — Graph assigns the "closest" content type instead. Pass the **list**
content type ID (the `…00<GUID>` form), not the site/base ID. Also note a Microsoft engineer's
statement that *"Microsoft Graph API for SharePoint doesn't have support for Document sets"* — Graph-created
Document Sets get no DocumentId until a later metadata edit.

Governing rule worth knowing: *"any content type that you assign to a document library must inherit
from Document… The exception to this pair of rules is the Folder content type and its derivatives."*
That is why a Document Set can live in a library despite not deriving from Document.

### D4 — Retention labels are never copied
**CONFIRMED · Effort: S · Risk: Low · Best value in this section**

`retentionLabel` is never read or written anywhere. `_ComplianceTag` /
`_ComplianceTagWrittenTime` are explicitly in the excluded `_builtInFields` list (~`:3495`). A retention
label on a folder is **inherited by everything placed in it**, so rebuilt folders and re-uploaded files
land unlabeled — a compliance-visible, silent failure.

Fix is cheap: `retentionLabel` is a first-class driveItem property with documented
`GET`/`PATCH /drives/{id}/items/{id}/retentionLabel`, explicitly supported on folders. Adding it to the
existing `$select` makes the read free.

### D5 — `remoteItem` is not detected
**CONFIRMED · Effort: S · Risk: Low**

`remoteItem` (an "Add shortcut to OneDrive" pointer into **another drive**) is not in the walk's
`$select`, so `item.Folder` may be null and the item is classified as a file — or, if a folder facet is
present, recursed into. Result: an empty folder at the target, a bogus zero-byte "file", or an
unexpected recursion into a foreign drive. Fix: add `remoteItem` to the `$select` and branch explicitly.
Free.

### D6 — ProgID repair may not flip the WOPI render flag
**PLAUSIBLE · Effort: S · Risk: Low**

Only two folder ProgIDs are attested anywhere: `OneNote.Notebook` (well corroborated) and
`SharePoint.DocumentSet` (community only). Microsoft documents no value list, and the premise that
ProgID was used for "document library folder types" appears unsupported — every ProgID in a real
shipped `DOCICON.XML` is a *file* ProgID.

The mechanism `PatchFolderProgIdAsync`'s comment doesn't mention: the writable surface is the folder's
list-item field **`HTML_x0020_File_x0020_Type`**, and setting it cascades to the internal ProgID **and
flips the WOPI render flag** so the folder opens in the OneNote web app. The current CSOM
`SetProperty(ProgID)` + `Folder.Update()` may set ProgID without flipping that flag — i.e. the notebook
still won't open correctly.

Bonus: this doubles as a **free second detection signal** — the walk already expands `listItem`, so
adding `fields($select=HTML_x0020_File_x0020_Type)` alongside `contentType` catches any special
container regardless of facet or content type, at zero extra round trips. A `.onetoc2` child is a third
signal.

### D7 — Folder-level default column values are silently lost, and can apply *wrong* metadata
**CONFIRMED · Effort: M · Risk: Low**

"Location-based metadata defaults" are a library-level setting keyed by folder path, stored in
`<library>/Forms/client_LocationBasedDefaults.html` as XML mapping absolute server-relative folder URLs
→ field → default value. The `Forms` folder is hidden and never copied. Worse than a silent omission:
defaults fire on `ItemAdded`, so uploads into a *target* that has its own defaults can acquire **wrong**
metadata. The hrefs are absolute source paths, so a naive file copy wouldn't work either.

Detection: one REST GET per library — `GET <lib>/Forms/client_locationbaseddefaults.html`; 404 means
none configured. Cheapest correct action is to warn.

### D8 — Private/shared Teams channel content is silently skipped
**CONFIRMED (documented) · Effort: M · Risk: Low**

Standard channels share the parent site's default library with a folder per channel — harmless, plain
`0x0120` folders. But **private and shared channels each get their own SharePoint site**, so their
content is in a separate drive and is silently skipped by walking the parent site's Documents library.
Detection needs channel enumeration, or at minimum a warning when the source site is Teams-connected.
Secondarily, a channel folder rebuilt at a target with no matching Teams channel is an orphan.

### D9 — Microsoft Loop
**CONFIRMED (documented) · Effort: S · Risk: Low**

`.loop`/`.fluid`/`.loot`/`.page`/`.pod` are **files**, not containers, so they copy — but Microsoft
documents that *"Loop components don't load if the file was moved to a different library"* and
*"Moving a `.loop` file from OneDrive to a SharePoint site results in the Loop component failing to
load."* Since a migration always writes to a different library, **every relocated Loop component breaks
its live embeds.** Whether a *copied* `.loop` opens standalone at the destination is unconfirmed (file
contents are documented as encrypted and not retrievable via Graph). Teams channel-created Loop content
lands in the ordinary channel folder, so this is genuinely reachable. Detection: extension match on
`name` — free. Action: warn rather than imply success.

**Loop workspaces** wrongly stored in an SPO site are a total loss: Microsoft's recognition criteria are
a `LoopAppData` folder in the default Documents library containing an `.appdata` subfolder with a
`.pod` file, on a site whose URL ends in a GUID. The only sanctioned migration is manual recreation in
the Loop UI; rebuilding file-by-file produces a dead folder tree. Detect the `LoopAppData` folder at a
library root (free) and refuse with an explanation. Normal Loop workspaces live in SharePoint Embedded
(`/contentstorage/CSP_<guid>/`) and are invisible to site/drive enumeration by design — worth stating
in the docs so nobody assumes coverage.

### D10 — File content types are lost in both engines
**CONFIRMED · Effort: M · Risk: Low**

`MigrationPackageBuilder.cs:575` hardcodes `ContentTypeId="0x0101"` on every `SPListItem`; Enhanced REST
never sets one; `"ContentType"` is in the excluded-fields list. So a Wiki Page (`0x010108`), Basic Page
(`0x010109`), Web Part Page (`0x01010901`), "Link to a Document" (`0x01010A`), or Master Page
(`0x010105`) all arrive as a generic Document. Worth addressing independently of the folder work.

### D11 — The hidden `Forms` folder
**UNRESOLVED — one-minute live test · Effort: S**

Every library has one, holding `template.xsn`, `client_LocationBasedDefaults.html`, custom view pages,
and Document Set welcome pages at `Forms/<DocSetContentTypeName>/docsethomepage.aspx`. **Whether Graph
`/children` returns it could not be confirmed** — strong expectation is that it does not, but if it *is*
returned, rebuilding it as a plain folder and re-uploading `.aspx` into it is a genuine hazard. Same
test covers picture-library `_t`/`_w` thumbnail folders.

Corollary regardless of the test result: a Document Set's welcome page is library-level config in
`Forms/`, so even a correctly content-typed Document Set at the target renders with the *target*
library's welcome page.

### D12 — In-place records
**PLAUSIBLE · Effort: M · Risk: Low**

Record declaration is a site-collection and list/library setting — there is **no folder-level** record
setting. Record status lives in hidden item fields `_vti_ItemDeclaredRecord` /
`_vti_ItemHoldRecordStatus` (community-attested only) with no write path, so re-uploaded records arrive
as ordinary documents. A "Block Edit and Delete" record also can't be modified or deleted at the source,
so source-side operations may fail. Modern tenants express this via retention labels instead — see D4.

### D13 — Confirmed non-issues
No action needed. **Slide Libraries** — discontinued, not available in SPO, and slides were single-slide
`.pptx` *files*, never containers. **InfoPath form libraries** — Forms Services retired July 14 2026;
folders are plain `0x0120`, template lives in `Forms/`. **Whiteboard** — `.whiteboard` are files in
OneDrive (*"SharePoint isn't yet supported"*), and Microsoft's own duplication method is an ordinary file
Copy-to. **Report libraries / Dashboards / PerformancePoint / Data Connection Libraries** — removed from
SharePoint Server SE, never an SPO service; `.odc`/`.udcx`/`.rdl` are plain XML files. **Records Center
"record series"** — a routing-table entry (Content Organizer rule), site-level config, not folder state.
**Wiki/Site Pages libraries** — all page types are Document-derived *files*; the libraries permit plain
`0x0120` folders (but see D10).

---

## E. Verified-correct — do not touch

Assessed during review and found correctly solved. Listed so nobody "fixes" them.

- **`$batch` metadata design** (~`:737-1157`): the sub-request-count arithmetic (10 items when versions
  are needed, 20 when not), the `multiVersionItemIds` skip, `versionsOnly` mode, and especially
  **surfacing 429s from inside a 200-OK `$batch` envelope** (~`:1027-1037`) — that last one is subtle
  and right.
- **`CreateThrottleAwareGate` + `RecentThrottleBackoff`** (~`:199-218`): cross-phase throttle-window
  inheritance is the right design. The gap is that C5's phase doesn't use it, not the mechanism.
- **`AdaptiveParallelismController`**: AIMD with grow-only-under-load is a well-reasoned detail — it
  prevents ramping to full width while idle and then bursting.
- **`TransferMemoryBudget`**: byte-denominated, FIFO, correct clamp for oversized single files, correct
  re-credit on the cancellation race. The right abstraction; it just needs to also cover Enhanced REST
  (B2).
- **SPMI large-file gate**: can no longer be bypassed (`MigrationJobService.cs:1509`, `:1536-1538`) —
  falls back to `job.SourceSize` and treats unknown sizes as large.
- **`MigrationPackageBuilder` encryption path** (~`:681-737`): `TryGetBuffer` to encrypt in place,
  one-shot `TryEncryptCbc` into a pre-sized array, `GetBuffer()` over `ToArray()`. Peak memory per
  version is near-minimal.
- **LOH compaction at batch boundaries** (`MigrationJobService.cs:1131-1141`), rate-limited to
  once/minute.
- **Range-based download resume** (~`:671-707` + `:2221-2285`), including the "server ignored my Range"
  200-vs-206 guard.
- **Chunked UI result flushing** (`CopyService.cs:120-130`) and `CopyResult`'s `BeginInvoke` override.
- **`IsSpecialContainer`** (~`:572-577`) replacing the per-folder ProgID probe. C6/D3 extend the same
  idea rather than revisiting it.
- **`_folderSegmentTasks` `Lazy<Task<T>>` dedup with no-cache-on-fault** (~`:3312-3346`).
- **`SendSharePointRequestAsync` takes a request FACTORY** (`Func<string, HttpRequestMessage>`) and
  builds a fresh message inside the retry loop — **no request-reuse bug**. The 401-refresh and
  429-backoff paths both dispose before `continue`.
- **Gate discipline**: every `WaitAsync` in `SharePointService` has a matching `Release()` in a
  `finally`.
- **`Throttled` handler lifetime**: all in-file subscriptions unsubscribe in a `finally` that runs
  before gate disposal; `HookThrottleTracker` is idempotent; the 8 external subscribe sites all pair up.
  No leak found.
- **CSOM `ProcessQuery` first-element-only `ErrorInfo` check** is correct by design — ProcessQuery aborts
  at the first failing action and reports it in element 0. (Minor: all three parsers only inspect
  `ErrorInfo` when the root is a JSON array; a 200 with a non-array body returns success.)
- **Date writes**: `ToUniversalTime()` + `InvariantCulture` everywhere. (Reads: see B12.)
- **No `async void`**, no `.Result`/`.Wait()` deadlock risk in `SharePointService`.
- **All-skip re-runs already skip the Enhanced REST folder pass** (`CopyService.cs:494`, gated on
  `anyFileCopied`) — so C5's cost lands on real-work runs, not steady-state ones. Note this is the same
  gate that causes B1; fix B1 without losing this.

---

## F. Suggested sequencing

**Batch 1 — cheap, low-risk, independently shippable. ✅ DONE.** C3 (digest + EnsureUser caching), C4
(`AddPermissionRow`), C8 (`$top=1&$select=id`), C9 (`Parallel.ForEachAsync`), plus the gate half of C5.
Together a large cut to the folder-correction and version-replay paths.

**Batch 2 — the truth-telling fixes. ✅ DONE**, plus A8, A9, A12, A13, B10, B12 folded in from later
passes. A1, A3, A4, A5, A6, A10, A11, B1, B13. All small and self-contained; each one currently
misinforms the user.

**Batch 3 — the big perf win. ✅ DONE.** C1. Largest single improvement (~20× on the re-run users hit
repeatedly), reuses methods already proven in the SPMI path, depended on nothing else here.

**Batch 4 — the latent crash. ✅ DONE.** B2. Not just slowness, and the documented >2 GB workaround
currently cannot work.

**Batch 5 — one coordinated scan change. NOT STARTED.** C2, C6, C5's re-walk removal, and A2 together.
They all hinge on the scan emitting folder items and carrying a wider `$select`; doing them at once
avoids touching the scan contract three times. **Resolve C2's `$expand=listItem`-on-`/delta` constraint
first.** Deliberately deferred — see the status note at the top of this document for why.

**Batch 6 — folder fidelity. NOT STARTED.** D2 (test first), then D4, D5, D3, D6, D1. Out of scope for
the July 26 pass (new-feature work, not bug-fixing).

**Live-tenant tests to run before or alongside the above:** D2 (does `/copy` preserve Document
Set-ness?), D11 (does `/children` return `Forms`?), the folder-color field names and whether
`stampcolor` works (see below), and whether `_ColorHex` values reproduce the same *visible* colour
across tenants if they are palette indices.

---

## G. Folder-color feature — status

Implemented 2026-07-25/26. Reads and writes are in place; **nothing has been verified against a live
tenant.**

Facts established during review:
- `_ColorHex` (`{3BDAB9AC-9E5D-44D4-BDE9-13B37E170618}`, Hidden) holds the value: `Type="Text"` holding
  a palette **index** as a numeric string `"0".."15"` (0/empty = yellow, 1 = dark red, … 15 = light
  pink), *not* an RGB value. `_ColorTag` (`{76D13CD2-1BAE-45A5-8B74-545B87B65037}`) is a related column.
  The underlying store is the folder property bag key `vti_colorhex`.
- **Both fields are `ReadOnlyField="TRUE"` on current SharePoint Online.** They can be read but not
  written as list fields; `cli-microsoft365` moved off that approach for exactly this reason.
- OData forbids a leading `_`, so REST/Graph `$select` and JSON payloads use `OData__ColorHex` /
  `OData__ColorTag`; `ValidateUpdateListItem`'s `FieldName` and CSOM `SetFieldValue` take the raw
  internal name. Both forms are needed.
- A SharePoint REST `$select` **hard-fails** on an unknown column; Graph's `fieldValueSet` is an OData
  **open type**, so an unknown name inside `fields($select=…)` is silently omitted instead.
- The supported write path is `POST /_api/foldercoloring/stampcolor(DecodedUrl=@a1)` with
  `{"coloringInformation":{"ColorHex":"8"}}`.

Current implementation:
- Read: `GetFolderColorValues` off an expanded `listItem.fields`, gated behind
  `GetFileMetadataAsync(..., includeFolderColor: true)` — **false by default**, so the per-file hot path
  is unaffected; only the three folder call sites opt in. Best-effort in a try/catch with a fallback to
  the plain request.
- Read (fast-path skip check): `GetFolderCurrentMetadataAsync` requests the `OData__`-prefixed names and
  **retries without them** on failure, so a bad column name can never disable the folder fast-path
  optimization.
- Write: `StampFolderColorAsync` via the `foldercoloring` endpoint, called **before** the
  date/author correction in both engines — the endpoint writes the folder's list item and therefore
  stamps Editor/Modified, so the correction that follows overwrites those side effects. Getting this
  order wrong was the original bug: the first version issued a second CSOM `ProcessQuery` with a plain
  `ListItem.Update()` *after* the correction, reverting Author/Editor on every folder while the log
  still reported success.
- Color is deliberately absent from both the CSOM batch and the `ValidateUpdateListItem` call — the
  fields are read-only, ProcessQuery aborts at the first failing action, and
  `ValidateUpdateListItem`'s commit is all-or-nothing.
- Failures are counted and logged separately from metadata failures so an unsupported tenant does not
  inflate "could not correct metadata".

Open items: confirm `stampcolor` actually applies on the target tenant (read the colour back in the
browser, not just via API); confirm `_ColorTag` is not *also* needed; decide whether a source colour
that was *removed* should clear the target's colour (currently it does not, consistent with how every
other field is treated); and note that a palette index may not render as the same visible colour across
tenants with different themes.
