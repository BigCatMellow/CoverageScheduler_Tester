# Coverage Scheduler

Coverage Scheduler is a Google Sheets + Google Apps Script tool for planning staff coverage when teachers or other scheduled staff are absent.

The main interface is a full-page Google Apps Script web app. The spreadsheet remains the source of operational data, while the web app gives office staff a simpler way to enter absences, manage coverage people, generate assignments, review conflicts, save the final plan, and create handouts.

> **Start here if you are using the scheduler day to day.** You usually do not need to edit the underlying sheets directly.

## Quick navigation

- [What the system does](#what-the-system-does)
- [The normal daily workflow](#the-normal-daily-workflow)
- [Common examples](#common-examples)
- [How coverage staff work](#how-coverage-staff-work)
- [How the scheduler decides](#how-the-scheduler-decides)
- [Understanding the workbook](#understanding-the-workbook)
- [First-time setup](#first-time-setup)
- [Updating the live web app](#updating-the-live-web-app)
- [Troubleshooting](#troubleshooting)
- [Repository layout](#repository-layout)

## What the system does

Coverage Scheduler combines five kinds of information:

1. **Who is scheduled to be where** — from `Teacher Schedule`.
2. **Who is absent and when** — entered through the web app and stored in `Daily Absences`.
3. **Which field trips change the normal day** — stored as editable events in `Field Trips`.
4. **Who can provide coverage** — managed through **+ Coverage Staff**, plus temporary teachers released by a field trip.
5. **The rules and limits for assigning coverage** — availability windows, tiers, restrictions, workload limits, and scheduler configuration.

When you click **Generate Plan**, the scheduler compares all of those pieces and builds a proposed coverage plan. It tries to avoid time conflicts, keep coverage practical, and use higher-priority coverage people before lower-priority options.

Generating a plan is a preview step. You can review and adjust the result before saving it to `Coverage Output`.

### Main things you can do

- Add a single full-day or partial-day absence.
- Create and edit **Field Trips** with the staff going, student grade(s), date, and time.
- View field trips and ordinary absences together in the monthly **Calendar**.
- Add or edit substitutes, aides, teachers with free blocks, administrators, or other coverage people.
- Mark coverage people available or unavailable for the selected day.
- Limit who can cover certain grades, subjects, assignment types, or time windows.
- Generate a coverage plan.
- View the plan as a timeline, table, or grouped by coverage person.
- Manually reassign individual blocks when needed.
- Save the approved plan to `Coverage Output` and create the printable Google Docs handout as separate actions.

## The normal daily workflow

Most days should follow this sequence:

### 1. Choose the date

Open the deployed web app and select the day you are planning.

The scheduler loads:

- that day's absences;
- the people available to provide coverage;
- the relevant teacher schedules;
- any previously saved plan data.

### 2. Enter absences

For one person, click **+ Add Absence**.

You can choose:

- **Full Day** — the person is unavailable for all scheduled blocks.
- **Partial Day** — enter the time they will be gone.
- **Emergency** — records the absence as an emergency so emergency fallback behavior can be considered by the scheduler.

You can also add notes and, for a normal single-person absence, optionally specify preferred coverage.

### Field trips on the selected day

Field trips are separate from ordinary absences. Use **+ Field Trip** to create one, or **Calendar** to find and edit an existing event.

A field trip is stored once rather than as several independent absence rows. It can be one day or span multiple/overnight dates. The trip name, destination, selected grade(s), staff, departure, and return determine both the field-trip coverage form and which classes disappear, which classes still need coverage, and which teachers staying behind become temporarily available.

### 3. Confirm the coverage team

Look at the **Coverage Staff** panel on the right.

Use the availability switch to confirm who is actually available that day. If someone needs to be added or edited, use **+ Coverage Staff**.

### 4. Generate the plan

Click **Generate**.

The scheduler finds the blocks that actually need coverage and attempts to assign available people without overlapping their assignments.

### 5. Review the result

Use whichever view is easiest:

- **Timeline** — useful for seeing the day chronologically.
- **Table** — useful for checking every block.
- **By Sub** — useful for seeing one coverage person's workload.

Field-trip rows use these same views and controls. When a day includes field-trip coverage, those rows are grouped under **Field Trip · <event name>** so the source is obvious without creating a second coverage interface.

Pay special attention to anything marked **Unfilled**.

### 6. Make manual adjustments if needed

If the automatically generated plan is not the arrangement you want, open an assignment and select a different coverage person.

After manual changes, review the affected person's other assignments to make sure the final plan still makes operational sense.

### 7. Save the final plan and create the handout

When the plan looks right, use the two actions independently:

1. Click **Save Plan** to write the exact plan you are looking at to `Coverage Output`.
2. Click **Create Handout** to build a Google Docs handout from the exact plan currently on screen and open it from the confirmation link.

Saving does not create a handout, and creating a handout does not save the plan. Manual reassignments made after generation are still preserved because both actions use the reviewed on-screen plan. The handout is organized by coverage person so each person can see where they need to be and when. Every generated handout is stored in a persistent Drive folder named from the workbook, such as `2627 Coverage Scheduler - Handouts`, so that folder can be shared once with anyone who needs access.

Field-trip rows use a generated version of the approved field-trip coverage form instead of the ordinary landscape handout. The scheduler creates a fresh Google Doc and builds only the forms actually needed, using the reference form's page geometry, typography, fields, table columns, and two-forms-per-page layout. The form's `WITH` column is filled only when Teacher Schedule identifies one unambiguous other teacher teaching the same class during that block; otherwise it is left blank. The old `Config!Field_Trip_Form_Template_ID` value is retained for visual-reference/backward-compatibility purposes but is no longer opened or copied by the runtime.

If there are no assigned coverage rows, the plan can still be saved, but there is nothing to include in a coverage-person handout.

## Common examples

### Example 1: One teacher is out all day

**Situation:** Jordan Lee is absent for the entire day.

1. Choose the correct date.
2. Click **+ Add Absence**.
3. Select **Jordan Lee**.
4. Choose **Full Day**.
5. Save the absence.
6. Confirm which coverage staff are available.
7. Click **Generate**.

The scheduler checks Jordan's schedule and only creates coverage needs for blocks that are marked or inferred as requiring coverage. Planning, break, or other non-cover-required blocks are normally ignored.

### Example 2: A teacher leaves for an appointment

**Situation:** Morgan Patel will leave at 11:15 AM and return at 1:45 PM.

1. Click **+ Add Absence**.
2. Select **Morgan Patel**.
3. Choose **Partial Day**.
4. Enter **11:15 AM** through **1:45 PM**.
5. Save and generate the plan.

The scheduler covers blocks that overlap that absence window. A class from 10:30–11:20 overlaps the absence and can require coverage; a class ending before 11:15 does not.

### Example 3: A 2nd-grade field trip changes both coverage needs and available teachers

**Situation:** Mike is going on a 2nd-grade field trip from 9:00 AM until 2:00 PM. Mike teaches both 2nd grade and 1st grade. Steve is staying at school and normally teaches 2nd grade during part of that window.

1. Click **+ Field Trip**.
2. Name the event, such as `2nd Grade Field Trip`.
3. Choose the departure date/time and return date/time. For a one-day trip, use the same date twice.
4. Select **2** under **Students on Trip**.
5. Select Mike and every other staff member going on the trip.
6. Save the field trip.
7. Click **Generate Plan**.

The scheduler sends the field-trip staff through the same block-by-block coverage-plan pipeline used for ordinary absences, with one important exception: classes for the students who are away are removed first.

That means:

- Mike's 2nd-grade classes during the trip are treated as cancelled because those students are away.
- Mike's 1st-grade classes still happen and become normal coverage-plan rows.
- Steve's 2nd-grade class during the trip is also treated as cancelled, so that block becomes temporary coverage availability.
- Only staff who teach the trip grade and are staying at school enter the field-trip coverage pool.
- Steve is preferred for Mike's overlapping 1st-grade coverage before an unrelated substitute when the times fit.
- A cancelled 2nd-grade class can be used directly for coverage.
- A planning period can be used when the existing schedule marks it cover-eligible.
- Availability can be assembled across adjacent blocks. For example, a cancelled 2nd-grade class from 12:00–1:15 plus a 1:15–1:30 break can cover one 12:45–1:30 need.
- If any of that coverage consumes break time, the scheduler must reserve the same amount of replacement break time inside another cancelled trip-grade class. That replacement time is then blocked from further coverage assignments.
- Every field-trip assignment records why the teacher is available, including mixed cases such as **2C class cancelled through 1:15; break time 1:15–1:30 moved to 1:30–1:45 inside cancelled 2D**.
- Steve's other classes that are still happening keep him unavailable.
- If the field-trip pool cannot cover a block, the scheduler falls back to the normal Coverage Staff pool.
- The finished result appears in the normal coverage plan with a specific person assigned to each remaining class, or **Unfilled** if no eligible person exists.

Field trips have stable event IDs and can be edited as one event from **+ Field Trip**, the left-side field-trip entry, or the **Calendar**. After generation, field-trip coverage is not shown in a separate planner. Its rows appear in the normal Timeline, Table, and By Sub views under a lightweight **Field Trip** group labeled with the event name, grades, and active trip window. Those rows use the same click-to-reassign behavior as every other coverage row.

### Manual placement and absence safeguards

Every generated coverage row can be manually reassigned through the normal coverage-plan row interaction, including rows grouped under a field trip.

The manual picker is calculated from the live workbook for that exact block. It may include:
- teachers released because the trip cancelled their trip-grade class;
- active Coverage Staff;
- other staff whose Teacher Schedule explicitly makes them cover-eligible at that time.

Configured substitutes are shown in a separate **Subs** group at the bottom of the picker. They still have to pass the same live availability, absence, schedule-conflict, and workload checks before they appear.

It excludes anyone who is absent during any portion of the block, on the field trip, already covering another overlapping block, teaching/duty-bound at that time, outside an availability window, or explicitly marked unavailable in Coverage Staff.

Manual placement does **not** bypass those safeguards. The server validates the selected person again when **Save Assignment** is pressed. Full-day and partial-day absences are both enforced, and multiple absence windows for the same person are retained instead of overwriting one another.

The older **Preferred Coverage** setting on an absence follows the same rule: it is treated as a preference only when that person can validly cover the required blocks. Otherwise the normal scheduler takes over.

### Field-trip assignment priorities

Field-trip coverage is scheduled across the **whole event in chronological order**, not one absent teacher at a time. This lets the scheduler make the same kind of continuity decisions a person would make when looking at the day as a whole.

For each block, the scheduler prefers:

1. a staying teacher whose trip-grade class is cancelled during that exact time;
2. among equally valid choices, a teacher whose released trip-grade time continues into upcoming coverage blocks;
3. reusing someone already helping with the same field trip;
4. keeping the same coverage person across adjacent blocks, even when the absent teacher changes;
5. direct released-class coverage over planning or a moved break;
6. if break time must be consumed, moving it to the nearest suitable cancelled trip-grade block, preferring a later block;
7. regular Coverage Staff only after trip-created coverage and ordinary absence scheduling are exhausted.

The goal is not merely to fill every block. It is to use **fewer people, fewer handoffs, and less schedule disruption** while preserving each teacher's required break time.

### Schedule term and trip-boundary rules

The native Teacher Schedule may contain both `S1` and `S2` rows for the same teacher. Coverage Scheduler filters to the semester that applies to the selected date before it decides whether someone is free. By default July–December uses **S1** and January–June uses **S2**. A Config value named `Schedule_Term_Override`, `Active_Term`, or `Schedule_Term` can override that inference when needed.

Partial-day coverage is clipped to the actual absence or trip window. If a trip begins at 9:00 AM during an 8:30–9:15 class, the coverage need is **9:00–9:15**, not the full 8:30–9:15 block.

### Overnight and multi-day field trips

A field trip can span several days without creating separate events.

For example, a trip that departs Monday at 10:00 AM and returns Wednesday at 1:00 PM is interpreted as:

- **Monday:** trip rules apply from 10:00 AM onward;
- **Tuesday:** trip rules apply for the full school day;
- **Wednesday:** trip rules apply until 1:00 PM;
- classes before departure Monday and after return Wednesday operate normally.

The same event ID appears across all affected dates in the Calendar, but editing any occurrence edits the single underlying field trip.

### Example 4: A substitute is only available in the morning

**Situation:** Casey Nguyen can cover from 8:00 AM until noon.

Edit Casey under **Coverage Staff** and set the normal or date-specific availability window appropriately.

The scheduler will consider Casey only for blocks that fit inside that availability window.

### Example 5: A teacher can cover only during free blocks

**Situation:** Taylor Brooks is a teacher who can occasionally cover classes during planning periods.

When adding Taylor as Coverage Staff:

- choose **Teacher (Free Blocks)** as the role;
- turn off **Can cover any block during these hours**;
- make sure Taylor also exists correctly in `Teacher Schedule`.

The scheduler uses Taylor's Teacher Schedule to identify blocks marked or inferred as coverage-eligible, such as planning periods.

## How coverage staff work

Coverage Staff are the people the scheduler may assign to an absent person's blocks.

Typical examples include:

- substitutes;
- aides;
- administrators;
- specialists;
- teachers who can cover during free blocks;
- other staff who are allowed to provide coverage.

### Priority tiers

The web app uses three priority tiers:

| Tier | Meaning | Typical use |
| --- | --- | --- |
| **1** | First choice | Regular substitute or primary coverage staff |
| **2** | Next choice | Secondary coverage option |
| **3** | Last resort | Emergency or least-preferred option |

Tier is one factor, not the only factor. Someone still has to be available and eligible for the specific block.

### Availability

Coverage staff can have:

- normal available weekdays;
- normal start and end times;
- daily availability overrides;
- an active/inactive status for the selected day.

### Optional restrictions

A coverage person can also be limited by:

- allowed grades;
- allowed subjects;
- allowed assignment types;
- maximum blocks per day;
- maximum different teachers per day;
- whether their day can be split across multiple absent teachers.

Leaving an optional restriction blank generally means there is no restriction for that field.

## How the scheduler decides

The scheduler is a practical heuristic, not a mathematical optimizer. It makes a sequence of reasonable assignment decisions rather than testing every possible combination.

In broad terms it:

1. Reads the selected day's ordinary absences and field trips.
2. Removes field-trip-grade classes that will not happen because those students are away.
3. Finds the remaining blocks that genuinely need coverage.
4. Finds teachers staying behind whose field-trip-grade classes were cancelled and temporarily adds them to that event's coverage pool.
5. Allows those affected teachers to use released class blocks and usable planning/break blocks during the trip.
6. Prefers the event-specific field-trip pool for field-trip-created coverage needs before falling back to normal Coverage Staff.
7. Keeps teachers unavailable while they are teaching any class that is still happening.
8. Applies partial-day time windows and normal availability restrictions.
9. Prevents a coverage person from being assigned to overlapping blocks.
10. Prevents someone who is themselves absent or on the field trip from being used as coverage during that window.
11. Honors a manually preferred assignment for a normal single-person absence when one was explicitly selected.
12. Uses whole-day continuity for ordinary absences when configured, but schedules field-trip needs block-by-block first so released capacity is not wasted.
13. Falls back to split coverage and then the normal coverage pool when needed.
14. Leaves a block **Unfilled** when no valid automatic candidate is available.

An **Unfilled** result is intentional information. It means the scheduler could not find a candidate that satisfied its current rules; it does not silently invent availability.

For a deeper explanation, see [`docs/wiki/How-the-Scheduler-Works.md`](docs/wiki/How-the-Scheduler-Works.md) and [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md).

## Understanding the workbook

The web app is the normal operating interface, but all live data is stored in the connected Google Sheet.

| Sheet | What it is for | Usually edited by |
| --- | --- | --- |
| `Teacher Schedule` | Source schedule for teachers/staff | Existing schedule process / administrator |
| `Staff List` | Stable roster shown in absence pickers | Setup or administrator |
| `Coverage Staff` | Coverage people and assignment rules | Web app |
| `Substitute Availability` | Date-specific availability overrides | Web app / scheduler workflow |
| `Daily Absences` | Ordinary absences for specific dates | Web app |
| `Field Trips` | One editable record per field trip: event ID, start/end dates and times, grades, staff, notes | Web app / Calendar |
| `Field Trip Coverage Pool` | Visible materialized candidate matrix for field-trip coverage blocks; operators may adjust `Enabled` and `Priority_Adjustment` | Scheduler / front office |
| `Coverage Output` | Final saved coverage assignments, including the linked field-trip Event ID when applicable | Web app |
| `Config` | Scheduler behavior settings | Advanced/admin use |
| `Lists` | Validation/helper values | Setup routine |
| `_Preview` | Temporary generated preview data | Scheduler |

`Lists` and `_Preview` are helper sheets and are normally hidden. `Field Trip Coverage Pool` stays visible so the front office can inspect and adjust the precomputed candidate pool.

### Teacher Schedule

The preferred source format is:

~~~text
Teacher | Term | Day | Start | End | Class | Subject | Room
~~~

Each row represents one block of one person's day.

`Start` and `End` are authoritative. The scheduler does not require your school to use fixed numbered periods.

The scheduler also understands several alternate header names. Setup/validation can add or work with the more detailed internal fields used for role, grade, assignment type, coverage requirement, and cover eligibility.

### Staff List

The absence picker comes from `Staff List`.

The simplest format is:

~~~text
Teacher
Jordan Lee
Morgan Patel
Taylor Brooks
~~~

If `Staff List` is missing or empty when setup runs, the system seeds it from unique names in `Teacher Schedule`. If it already contains names, setup preserves it.

## First-time setup

For the full beginner walkthrough, start with [`docs/wiki/Home.md`](docs/wiki/Home.md).

Recommended order:

1. [`Getting Started`](docs/wiki/Getting-Started.md)
2. [`Install Google Apps Script`](docs/wiki/Installing-Google-Apps-Script.md)
3. [`Set Up the Workbook`](docs/wiki/Workbook-Setup.md)
4. [`Teacher Schedule`](docs/wiki/Teacher-Schedule.md)
5. [`Staff List`](docs/wiki/Staff-List.md)
6. [`Coverage Staff`](docs/wiki/Coverage-Staff.md)
7. [`Deploy the Web App`](docs/wiki/Deploying-the-Web-App.md)
8. [`Daily Workflow`](docs/wiki/Daily-Workflow.md)
9. [`Troubleshooting`](docs/wiki/Troubleshooting.md)

### Short installation version

1. Open the Google Sheet that will hold the coverage system.
2. Make sure the operational schedule is in a tab named **Teacher Schedule**.
3. Open **Extensions → Apps Script**.
4. Create matching Apps Script files and copy the files from `google-apps-script/`.
5. Include `appsscript.json` as the project manifest.
6. Save the Apps Script project and reload the spreadsheet.
7. Use **Coverage Scheduler → Set up workbook**.
8. Use **Coverage Scheduler → Validate teacher schedule**.
9. Deploy the Apps Script project as a **Web app**.
10. Open the generated `/exec` URL.
11. Add your real coverage team with **+ Coverage Staff**.

### What setup does

Running **Set up workbook**:

- creates the scheduler-managed sheets if they do not already exist;
- adds required headers without intentionally replacing existing populated schedule data;
- seeds helper lists and default configuration;
- creates/seeds `Staff List` when needed;
- hides helper sheets;
- stores the connected spreadsheet ID in Apps Script **Script Properties** so the standalone web app knows which workbook to reopen.

The repository does not need a hard-coded operational spreadsheet ID.

## Updating the live web app

Changes merged into GitHub do **not** automatically appear in the deployed Google Apps Script web app.

To update an existing installation:

1. Copy the changed files from `google-apps-script/` into the existing Apps Script project.
2. For changes that span the browser and server, update the complete related set **before** deploying. In particular, manual placement and current field-trip scheduling require the current `code.gs`, `scheduler.gs`, `index.html`, and `field-trip-ui.gs` together.
3. Save the project.
4. Choose **Deploy → Manage deployments**.
5. Edit the existing Web App deployment.
6. Select **New version**.
7. Deploy.

The web app carries an API compatibility version. If the UI and server are from different releases, it now stops at startup with an explicit **deployment is out of sync** message instead of allowing partially working behavior.

Using the existing deployment keeps the same `/exec` URL.

### Important distinction

- **GitHub `main`** = current source code.
- **Apps Script editor** = the code currently installed in your Google project.
- **Deployed Web App version** = the code users actually receive from the `/exec` URL.

All three need to be current for a code change to be visible to users.

## Troubleshooting

### I added an absence but no blocks appear

Check:

- the selected date/day matches the person's `Teacher Schedule` rows;
- the teacher name matches the schedule name;
- the scheduled blocks have valid start/end times;
- the blocks are considered to require coverage;
- a partial-day absence actually overlaps the expected blocks.

### The scheduler leaves something Unfilled

Check whether:

- any coverage staff are active that day;
- their availability windows include the block;
- their grade/subject/type restrictions allow the assignment;
- they are already covering someone else at the same time;
- they have reached a configured daily limit;
- the potential coverage person is also absent during that time.

An Unfilled block may be correct if there genuinely is no eligible person.

### A coverage person does not appear as an option

Check their **Coverage Staff** record, weekday availability, active status, time window, and restrictions.

If they are a **Teacher (Free Blocks)**, also check that their own Teacher Schedule contains a genuinely cover-eligible free/planning block at that time.

### I changed GitHub but the web app still looks old

Merging GitHub code is not enough. Copy the updated files into Apps Script, then create a **New version** of the existing deployment.

If you see **runner[method] is not a function**, the deployed browser and server files are from different releases. Update `code.gs`, `scheduler.gs`, `index.html`, and `field-trip-ui.gs` together and deploy one new version. Current builds detect this mismatch at startup and show a clearer compatibility error.

### I copied the project to a different spreadsheet and it still opens the old workbook

Open the new spreadsheet and run **Coverage Scheduler → Set up workbook**. That updates the spreadsheet ID stored in Script Properties.

### The absence picker is missing someone

Check `Staff List`. The web app deliberately uses that stable roster rather than rebuilding the picker from the schedule every time.

## Repository layout

~~~text
google-apps-script/
  code.gs                       Web-app entry point, workbook binding, UI additions
  web-ui-data.gs                Staff List and Coverage Staff web actions
  handout.gs                    Google Docs coverage handouts
  field-trip-ui.gs              Field trip editor, calendar, and event UI
  index.html                    Main full-page web interface
  teacher-schedule-adapter.gs  Teacher Schedule preservation/validation
  setup.gs                      Workbook schemas, setup, validation, defaults
  scheduler.gs                  Coverage engine, preview, output logic
  sidebar.html                  Optional spreadsheet sidebar
  sidebarcss.html               Optional sidebar styles
  sidebarjs.html                Optional sidebar client behavior
  appsscript.json               Apps Script manifest

apps-script/
  Legacy/alternate spreadsheet-sidebar implementation files

docs/
  wiki/                         Beginner and operator guides
  ARCHITECTURE.md               Technical data flow and design
  TEACHER-SCHEDULE.md           Schedule format and inference rules
  REVIEW.md                     Review findings / technical debt notes

prototype/
  Browser-only interface experiments and mock-data prototypes
~~~

The `google-apps-script/` folder is the authoritative copy-ready implementation for the current full-page web app.

## Configuration

`Config` is created with defaults such as:

- try whole-day coverage first;
- allow split coverage;
- fallback block/teacher limits;
- whether lunch may be used as cover-eligible time;
- date-based availability override behavior;
- the Apps Script time zone.

Most day-to-day users should not need to edit `Config`.

## Performance

The current Apps Script implementation keeps Google Sheets as the source of truth but avoids rebuilding field-trip candidate eligibility on every Generate. The visible `Field Trip Coverage Pool` materializes one row per eligible candidate per field-trip coverage block, including source, baseline score, and reason. Saving/editing a trip rebuilds its rows; Generate reads those rows first and then applies only live safety checks such as absences, current assignments, schedule conflicts, and daily limits. Operators can set `Enabled` to No or use `Priority_Adjustment` to steer ranking without bypassing hard constraints.

The main web bootstrap is also lightweight: it loads the Staff List without serializing every teacher block. A selected teacher's detailed schedule is fetched only when the user opens that person's schedule controls, then cached in the browser for that date.

Every web endpoint emits a timing/counter summary to Apps Script execution logs without logging staff names, absence details, or schedule contents. See [`docs/PERFORMANCE.md`](docs/PERFORMANCE.md) for the cache/invalidation model and what to measure after deployment.

No external database, third-party cache, new API dependency, or additional credential is required for these optimizations.

## Privacy and repository safety

Do **not** commit live school data to this repository.

Keep these outside version control:

- real staff schedules if they are operational/private;
- absence records;
- substitute availability containing private information;
- student information;
- spreadsheet exports;
- credentials or tokens;
- local clasp configuration containing sensitive project identifiers.

If sample data is ever added, deliberately anonymize it.

See [`SECURITY.md`](SECURITY.md).

## Current limitations

- There is no automated test suite yet.
- There is no CI deployment workflow yet.
- The Apps Script project is updated/deployed manually.
- The scheduling algorithm is heuristic and can produce a valid but non-optimal plan when many constraints compete.
- `Term` is currently mostly informational when schedules use `All Year`; seasonal schedules need explicit date-to-term mapping.
- Numeric grade inference works best when class labels contain recognizable grade numbers.
- Unusual PreK/Kindergarten/Beginner labels may need explicit grade configuration when strict grade restrictions are used.

## License

No open-source license is currently included. Unless a license is added, normal copyright restrictions apply.
