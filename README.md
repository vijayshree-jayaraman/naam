# naam: One-stop package for consistent string encoding, ID management, and label tracking across Stata datasets

`naam` ("What's in a naam?") solves a core problem in large-scale survey and administrative data work: encoding string variables *consistently* across multiple files.

Stata's built-in `encode` assigns numeric codes alphabetically within each dataset independently. If a later file introduces a new category (a new region, district, or industry), all alphabetically subsequent codes shift, and any merge or append across files produces wrong results with no error message.

`naam` encodes once, saves the exact string-to-numeric mappings, and reapplies them instantly to every subsequent file. The same string always gets the same number.

---

## Installation

Install the released version from SSC:

```stata
ssc install naam, all
```

Install the development version from GitHub:

```stata
net install naam, from("https://raw.githubusercontent.com/vijayshree-jayaraman/naam/main/")
```

---

## Subcommands

| Subcommand | What it does |
|---|---|
| `naam encode` | Encode string variables and save mappings to Excel |
| `naam apply` | Reapply saved mappings to a new file; auto-assigns new categories |
| `naam id` | Convert string IDs to consistent numerics; saves mappings as .dta |
| `naam id import` | Convert numeric IDs back to original strings using saved .dta mapping |
| `naam export` | Save value labels from an already-encoded dataset to Excel |
| `naam list` | Inspect a mapping file from inside Stata |
| `naam decode` | Reverse encoding: numeric back to original strings (for Excel-mapped vars) |
| `naam check` | Compare in-memory labels against saved mapping (QA) |
| `naam compare` | Compare two mapping files against each other |

---

## Quick start

```stata
* Round 1: encode and save mappings
use naam_round1.dta, clear
naam encode district occupation religion using naam_maps.xlsx, replace
naam id hhid using naam_ids, replace keep
save round1_enc.dta, replace

* Round 2: apply same mappings (new categories detected automatically)
use naam_round2.dta, clear
naam apply using naam_maps.xlsx
naam id hhid using naam_ids, keep
append using round1_enc.dta

* Every district code is now consistent across both rounds
tab district

* Convert numeric hhid back to original string IDs
naam id import hhid using naam_ids
```

---

## naam id import

`naam id import` is the reverse of `naam id`. After running `naam id`, your string ID variable is stored as a numeric. This command reads the saved `.dta` mapping file and restores the original string values.

**Syntax:**
```stata
naam id import varlist using basepath [, keep suffix(str)]
```

**Options:**
- `keep` - retain the numeric variable as `_num_varname` instead of dropping it
- `suffix(str)` - create a new string variable named `varname+suffix` instead of replacing the original
- `keep` and `suffix()` may not be combined

**Examples:**
```stata
* Replace hhid with original string (default)
naam id import hhid using naam_ids

* Keep numeric as _num_hhid, restore string to hhid
naam id import hhid using naam_ids, keep

* Process multiple ID variables at once
naam id import hhid person_id lender_id using naam_ids
```

---

## Why naam id saves .dta, not Excel

`naam id` stores mappings as native Stata `.dta` files (`base_varname.dta`) rather than Excel sheets. ID variables in administrative data can easily exceed Excel's hard row limit of 1,048,576. Storing natively in `.dta` removes this limit entirely, speeds up lookup via Stata's merge engine, and keeps the mapping in a format that is already part of any Stata workflow.

In practice, switching from string IDs to numeric using `naam id` can reduce file sizes dramatically. On one administrative dataset with multiple ID columns, this brought a 16 GB file down to 5 GB.

---

## Changelog

### v1.1.1
- **Bug fix**: `naam id` now preserves the full prior ID mapping when processing later files, so numeric codes are never reused for all-new or partially-new rounds.
- **Bug fix**: `naam id` validates saved mappings before merge; exact duplicate rows are tolerated, but conflicting duplicate strings or duplicate numeric codes now stop the command.
- **Bug fix**: `naam id, strict` now fails if no valid saved mapping exists, and exits immediately when unexpected new IDs appear.
- **Bug fix**: `naam id import` validates mapping uniqueness before merge and refuses to replace a numeric ID with blank strings when unmatched IDs are present; use `,suffix()` to inspect unmatched values safely.
- **Robustness**: Excel mapping files now store a `sheetname` column in the index and abort on 31-character sheet-name collisions instead of silently overwriting a mapping sheet.
- **Docs**: clarified `,replace` behavior and documented the implemented `naam encode, all` option.

### v1.1.0
- **New**: `naam id import` subcommand - convert numeric IDs back to original strings using saved `.dta` mapping files
- **Bug fix**: `naam id` - duplicate rows were silently created when an existing mapping had duplicate `string_value` entries; now deduplicated before merge
- **Bug fix**: `naam id` - new ID codes for unmatched observations were grouped using the wrong variable (the already-renamed numeric version); fixed to use the string backup `_str_v`
- **Bug fix**: `naam id` - `sort` added before `egen group()` to ensure stable, alphabetical code assignment on fresh encode
- **Bug fix**: `naam id` - the `,replace` option was accepted in syntax but silently ignored; invalid existing mappings now require `,replace` before they are discarded
- **Bug fix**: `naam id import` - the `,suffix()` and `,keep` options previously produced the same output regardless of which was specified; redesigned so the three modes are distinct and mutually exclusive (see help file)
- **Bug fix**: `naam encode` / `naam export` / `naam apply` - `numeric_code` was stored as `str20` (string) in Excel sheets; this caused type mismatches on merge and label define. Now stored as `long` in the dataset before export
- **Bug fix**: `naam apply` / `naam decode` - `nval == 0` (empty mapping sheet) was not guarded, causing macro errors; now skipped with a warning
- **Bug fix**: `naam apply` - updated hint message for ID variables now correctly references `naam id import`
- **Bug fix**: `naam list` - ID variables now display a hint about `naam id import`
- **Bug fix**: `naam id` - `n_processed` was incremented at variable entry (after the type check) rather than at completion, so if all variables were skipped by the `,replace` guard the final "no variables found" error was never raised; counter now incremented only on successful completion
- **Bug fix**: `naam apply` - mapping sheet load only confirmed `numeric_code` existed, not `string_value`; a malformed sheet with only one column would load silently and produce empty labels with no error; both columns are now confirmed
- **Bug fix**: `naam check` - if the user supplied a varlist but every variable failed the `confirm variable` check, the command exited silently having checked nothing; now raises an explicit error
- **Bug fix**: `naam id` - duplicate `confirm file` call (once for the replace guard, once for `file_exists`) removed; single check now sets `file_exists` directly, eliminating the redundancy

### v1.0.1
- Initial release

---

## Requirements

Stata 14 or higher. No user-written dependencies.

---

## Citation

If you use `naam` in your work, please cite:

> Jayaraman, Vijayshree (2026). "naam: Consistent string encoding, ID management, and label tracking across Stata datasets." Available at: https://github.com/vijayshree-jayaraman/naam and https://ideas.repec.org/c/boc/bocode/naam.html

`naam` was inspired by `codebookout` (Das, 2014):

> Das, Kishor K. (2014). "CODEBOOKOUT: Stata module to save codebook in MS excel format." *Statistical Software Components* S457811, Boston College Department of Economics. https://ideas.repec.org/c/boc/bocode/s457811.html

---

## Author

Vijayshree Jayaraman
[jvijayshree26@gmail.com](mailto:jvijayshree26@gmail.com)
