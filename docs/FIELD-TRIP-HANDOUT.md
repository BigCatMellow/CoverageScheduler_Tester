# Field-Trip Handout Specification

The field-trip handout is generated directly by `google-apps-script/handout.gs`. The former Google Doc template is retained only as a human visual reference; runtime handout creation must not open, copy, append, trim, or mutate that template.

## Page geometry

- US Letter portrait: 612 × 792 points.
- Margins: 36 points (0.5 inch) on every side.
- Printable content width: 540 points.
- Two field-trip forms per page. After every second form, start a new page.
- When two forms share a page, preserve the reference form's compact separation with six blank 11-point lines between them.

## Form typography and labels

Metadata lines use Cambria, 12 point, bold, with 150% line spacing.

The form contains these labels:

1. `FIELD TRIP COVERAGE FORM` — underlined — and `PREPARED BY: C. SCHREMBS` on the same line.
2. `COVERAGE FOR:` and `DATE:`.
3. `CLASS(ES) TAKING TRIP:`.
4. `TRIP DESTINATION:`.
5. `DEPARTURE TIME:` and `APPROXIMATE RETURN TIME:`.

## Assignment table

Each form contains one five-column table using the full 540-point content width.

| Column | Width | Source |
| --- | ---: | --- |
| NAME | 105 pt | `Assigned_Coverage` |
| TIME NEEDED | 87 pt | formatted `Start–End` |
| RM#/SUBJECT | 98 pt | Room plus Subject/Class fallback |
| SPECIAL INSTRUCTIONS | 175 pt | field-trip notes |
| WITH | 75 pt | one unambiguous matching same-class teacher, otherwise blank |

Table text uses Cambria 11 point and centered paragraphs. The header row is bold with a 14-point minimum height. Assignment rows use a 15-point minimum height, and the NAME column is bold. Borders are black, 1 point.

## Form grouping

- Rows are grouped by the staff member who is on the trip (`Absent_Staff`).
- Rows inside each group are sorted chronologically, with class as a stable secondary sort.
- One form holds at most six assignments.
- More than six assignments for the same staff member continue on another form.
- Only required forms are created; there are no blank placeholder forms to remove.

## Safety boundary

Handout generation remains downstream of the existing live plan revalidation. Stale absences, field trips, schedules, conflicts, limits, or other invalid assignments must still be rejected before rendering.

The rendering change does not alter Save Plan semantics, scheduling eligibility, field-trip candidate selection, or manual reassignment validation.

## Legacy reference

`Config!Field_Trip_Form_Template_ID` may remain populated so the approved historic Google Doc can be consulted visually. It is not a runtime dependency of field-trip handout creation.
