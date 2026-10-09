# Google Apps Script Version

This folder is the copy-ready Google Apps Script version of Coverage Scheduler.

The primary interface is the **full-page web app** in `index.html`, based on the original Coverage Scheduler layout. The spreadsheet sidebar files are retained as an optional secondary interface.

For a complete first-time-user walkthrough, start with [`docs/wiki/Home.md`](../docs/wiki/Home.md).

## Files

- `code.gs` — web-app entry point, spreadsheet menu, workbook binding, and server wrappers
- `web-ui-data.gs` — Staff List roster adapter and coverage-team create/edit/remove functions
- `handout.gs` — full-width, print-friendly Google Docs coverage handouts
- `field-trip-ui.gs` — field trip editor, monthly calendar, and event UI
- `index.html` — thin full-page web-app shell
- `styles.html` — full-page interface styles
- `ui-core.html` — browser state, RPC helpers, schedule helpers, bootstrap, and base rendering
- `ui-plan.html` — Timeline/Table/By Sub plan rendering
- `ui-people.html` — absence and Coverage Staff editing workflows
- `ui-manual.html` — manual block reassignment UI
- `ui-actions.html` — generate/save/handout actions and event binding
- `teacher-schedule-adapter.gs` — preserves and validates the existing Teacher Schedule source
- `setup.gs` — managed workbook sheets, validation, and defaults
- `scheduler.gs` — top-level generation orchestration
- `coverage-foundation.gs` — shared schemas, cache plumbing, sheet I/O, config, and time/value utilities
- `field-trips.gs` — field-trip persistence, normalization, date behavior, and participant absences
- `schedule-cache.gs` — prepared weekday Teacher Schedule cache and schedule normalization
- `coverage-staff.gs` — Coverage Staff roster, date overrides, and eligibility normalization
- `field-trip-pool.gs` — field-trip candidate pool, freshness, released-time logic, and break relocation
- `manual-coverage.gs` — manual reassignment workspace and live assignment validation
- `coverage-assignment.gs` — coverage needs, scoring, automatic assignment, and plan-row construction
- `coverage-save.gs` — durable plan save/reload plus legacy preview helpers
- `sidebar.html` — optional spreadsheet sidebar markup
- `sidebarcss.html` — optional sidebar styles
- `sidebarjs.html` — optional sidebar browser logic
- `appsscript.json` — Apps Script project manifest

## Teacher Schedule and Staff List

The preferred `Teacher Schedule` tab is:

```text
Teacher | Term | Day | Start | End | Class | Subject | Room
```

The scheduler reads this sheet directly. `Start` and `End` are authoritative; fixed school periods are not required. Grade information is especially important for field trips because it determines which classes are cancelled and which staying teachers become temporarily available. Common forms such as `2`, `2nd Grade`, `Grade 2`, `PreK`, `Kindergarten`, and `Beg` are normalized by the scheduler.

The web app uses a separate `Staff List` tab as the roster shown in **+ Add Absence**. Its required header is simply:

```text
Teacher
```

If `Staff List` does not exist or is empty, setup creates it and seeds unique teacher names from `Teacher Schedule`. If it already contains names, the app leaves it alone and treats it as the roster source.

## Install

1. Open the Google Sheet you want to use.
2. Confirm the live schedule is in a tab named **Teacher Schedule**.
3. Go to **Extensions → Apps Script**.
4. Create matching files and copy in the contents from this folder:
   - `code.gs`
   - `web-ui-data.gs`
   - `handout.gs`
   - `field-trip-ui.gs`
   - `teacher-schedule-adapter.gs`
   - `setup.gs`
   - `coverage-foundation.gs`
   - `field-trips.gs`
   - `schedule-cache.gs`
   - `coverage-staff.gs`
   - `field-trip-pool.gs`
   - `manual-coverage.gs`
   - `coverage-assignment.gs`
   - `scheduler.gs`
   - `coverage-save.gs`
   - `index.html`
   - `styles.html`
   - `ui-core.html`
   - `ui-plan.html`
   - `ui-people.html`
   - `ui-manual.html`
   - `ui-actions.html`
   - `sidebar.html`
   - `sidebarcss.html`
   - `sidebarjs.html`
5. If the manifest is hidden, open **Project Settings** and enable **Show `appsscript.json` manifest file in editor**, then replace it with this folder's manifest.
6. Save the project and reload the spreadsheet.
7. Choose **Coverage Scheduler → Set up workbook**. This creates the scheduler tabs, creates/seeds `Staff List` when needed, and remembers this spreadsheet for the standalone web app.
8. Choose **Coverage Scheduler → Validate teacher schedule**.

You do **not** need to manually edit `Coverage Staff` for normal use. Open the web app and use **+ Coverage Staff** to add or edit the coverage team.

## Deploy the full-page interface

1. In the Apps Script editor choose **Deploy → New deployment**.
2. Choose **Web app**.
3. Choose the execution identity and access level appropriate for your organization.
4. Deploy and authorize the requested Google Sheets/Docs/Drive permissions.
5. Open the generated `/exec` URL.

For normal manual updates, **do not copy the modular source files one by one**. GitHub Actions now builds a copy-ready artifact named `coverage-scheduler-apps-script` after every push to `main`. Download that artifact and replace only `CoverageScheduler.gs`, `index.html`, and `appsscript.json` in Apps Script. See [APPS-SCRIPT-MANUAL-DEPLOY.md](../docs/APPS-SCRIPT-MANUAL-DEPLOY.md). The modular files in this folder remain the maintainable source of truth.

The web UI supports the normal workflow: choose a date, add/edit absences, create/edit/remove coverage staff, toggle daily coverage availability, generate the plan, inspect Timeline/Table/By Sub views, and manually reassign blocks. **Save Plan** persists the exact reviewed plan to `Coverage Output`. **Create Handout** now persists that same reviewed plan first and only then creates the handout, so a generated handout always has reloadable saved coverage behind it. Returning to that date reloads the durable saved rows from `Coverage Output`; unsaved generated plans remain browser-only and are intentionally not restored from `_Preview`.

Field trips are first-class events rather than ordinary group absences. Use **+ Field Trip** to enter the trip name and destination, then choose the departure date/time, return date/time, student grade(s), and staff going on the trip. For one-day trips, the start and end dates are the same. Overnight trips remain one event across the full date range. Use **Calendar** to see field trips and ordinary absences together and to reopen an event for editing.

Manual reassignment uses the same scheduling engine as automatic generation. The web app requests eligible choices for the selected block from the server, then validates the final selection again before changing the in-memory plan. Live daily absences, field-trip participant windows, existing plan assignments, Coverage Staff availability, and Teacher Schedule conflicts are all applied. Absence state supports multiple windows per person so one absence cannot overwrite another. Configured staff whose role is **Substitute** are presented in a separate **Subs** optgroup at the bottom of the assignment picker; grouping does not bypass any eligibility checks.

Field-trip coverage uses the **same plan renderer and reassignment UI** as ordinary coverage. Rows linked by `Event_ID` are grouped under a compact **Field Trip** section labeled with the event name, grades, and active trip window in Timeline, Table, and By Sub views. The field-trip editor/calendar remains separate because it defines the event; the resulting coverage rows do not use a separate mini-planner.

Field-trip needs from all trip participants are flattened into one chronological event timeline before assignment. Candidate scoring rewards direct released-class availability, future released-time runway, reuse within the same field trip, and adjacent-block continuity across different absent teachers. Break relocation prefers the nearest suitable released block after the consumed break window when possible. Regular Coverage Staff are deferred until after ordinary absence scheduling if the trip-created pool cannot solve a block.

During generation, trip-grade classes are treated as cancelled while those students are away. On a multi-day trip, the departure day begins at the departure time, middle dates are treated as full-day trip dates, and the return day ends at the return time. The field-trip coverage pool contains only staff who teach the trip grade and are staying at school. Teacher Schedule rows are filtered to the active semester for the selected date before availability is calculated. Their cancelled trip-grade classes can provide direct coverage; adjacent released/planning/break segments may be combined to cover one continuous need; and any break time consumed by coverage must be moved into reserved time inside another cancelled trip-grade class. The generated plan records the exact availability reason for each field-trip assignment. Partial-day needs are also clipped to the true event boundary, so a class that overlaps the start or end of a trip produces only the portion that actually needs coverage. Field-trip coverage rows carry the event ID so the normal coverage views can group them under the correct **Field Trip** heading while retaining the same Timeline/Table/By Sub interactions and manual reassignment flow. Staff on the trip still need coverage for their other classes. Teachers staying behind whose trip-grade classes were cancelled become a temporary first-priority coverage pool; their released class blocks and usable planning/break blocks can cover the classes that remain. The scheduler falls back to normal Coverage Staff only when the field-trip pool cannot cover a block.

## Handouts

**Save Plan** writes the reviewed plan to `Coverage Output`. **Create Handout** saves that same reviewed in-memory plan first and then builds the Google Doc, so every generated handout has a durable reloadable schedule behind it. Manual block reassignments made after generation are preserved in either action.

For ordinary absence rows, handouts keep the generated landscape format. Field-trip handouts are now built directly from the approved form specification in code instead of copying and trimming a Google Doc template. Each form is for one staff member who is on the trip and can hold up to six coverage assignments; two forms are placed on each page. The generated form preserves the reference document's Letter portrait page, 0.5-inch margins, Cambria typography, metadata fields, and five-column assignment table. The `WITH` cell is populated only when Teacher Schedule resolves exactly one other regular teacher for the same class/block; ambiguous or missing matches remain blank. `Config!Field_Trip_Form_Template_ID` is retained only as a legacy/reference value and is not opened during handout generation. See [`../docs/FIELD-TRIP-HANDOUT.md`](../docs/FIELD-TRIP-HANDOUT.md) for the construction specification.

On the first handout creation for a workbook, the app creates (or reuses) a sibling Drive folder named `<Workbook Name> - Handouts`. Its folder ID is remembered for that workbook, and every later handout is moved into that same folder. The success message includes links to both the new handout and the folder, so the folder can be shared once instead of sharing each document individually.

For ordinary absence coverage, `handout.gs` creates one landscape page per coverage person with a full-width assignment table. The time column is intentionally wide enough for normal time ranges to stay on one line, while compact cell padding keeps rows short. The table uses the full printable width of the page with larger Subject and Absent Teacher columns for easier scanning.

## Workbook binding

The public repository does **not** contain a hard-coded operational spreadsheet ID.

When you run **Coverage Scheduler → Set up workbook** from the target spreadsheet, `code.gs` stores that spreadsheet ID in Apps Script **Script Properties**. The standalone web app reopens that workbook on future requests.

If you copy the project to a different workbook, run **Set up workbook** from the new spreadsheet to update the stored connection.

## Performance behavior

The web app treats each server call as one request snapshot. Field-trip block eligibility is materialized into the visible `Field Trip Coverage Pool`: each eligible candidate/block combination is precomputed when a trip is saved or when the pool is rebuilt. Generate reads that pool and performs final live checks instead of scanning the full candidate universe again. Front-office users may set `Enabled` to No or edit `Priority_Adjustment`; these edits steer selection but do not bypass hard schedule/absence/limit validation. The initial web bootstrap still sends a lightweight Staff List, and detailed Teacher Schedule data is loaded only when needed.

Apps Script execution logs include a `[Coverage Performance]` JSON summary for each web endpoint. See [`../docs/PERFORMANCE.md`](../docs/PERFORMANCE.md) for details.

## Updating an existing web deployment

Saving newer code in Apps Script does not automatically update a versioned production deployment.

After copying updated files:

1. choose **Deploy → Manage deployments**;
2. edit the existing Web App deployment;
3. select **New version**;
4. deploy again.

The existing `/exec` URL can continue to be used.