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
- Fixed ID mapping updates so numeric IDs are never reused across files.
- Added stronger validation for duplicate/corrupt mappings and high-cardinality IDs.
- Improved `naam id import`, strict mode, row-order preservation, and documentation.

### v1.1.0
- Added `naam id import`.
- Fixed several mapping and label-restore edge cases.

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
