# naam — consistent string encoding, ID hashing, and label management across datasets

`naam` ("What's in a naam?") solves a core problem in large-scale survey and administrative data work: encoding string variables *consistently* across multiple files.

Stata's built-in `encode` assigns numeric codes alphabetically within each dataset independently. If a later file introduces a new category — a new region, district, or industry — all alphabetically subsequent codes shift, and any merge or append across files produces wrong results with no error message.

`naam` encodes once, saves the exact string-to-numeric mappings, and reapplies them instantly to every subsequent file. The same string always gets the same number.

---

## Installation

From SSC (recommended):

```stata
ssc install naam, all
```

From GitHub:

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
| `naam export` | Save value labels from an already-encoded dataset to Excel |
| `naam list` | Inspect a mapping file from inside Stata |
| `naam decode` | Reverse encoding — numeric back to original strings |
| `naam check` | Compare in-memory labels against saved mapping (QA) |
| `naam compare` | Compare two mapping files against each other |

---

## Quick start

```stata
* Round 1: encode and save mappings
use round1.dta, clear
naam encode district occupation religion using naam_maps.xlsx, replace
naam id hhid using naam_ids, replace keep
save round1_enc.dta, replace

* Round 2: apply same mappings (new categories detected automatically)
use round2.dta, clear
naam apply using naam_maps.xlsx
naam id hhid using naam_ids, keep
append using round1_enc.dta

* Every district code is now consistent across both rounds
tab district
```

---

## Why naam id saves .dta, not Excel

`naam id` stores mappings as native Stata `.dta` files (`base_varname.dta`) rather than Excel sheets. This is because ID variables in administrative data can easily exceed Excel's hard row limit of 1,048,576. Storing natively in `.dta` removes this limit entirely, speeds up lookup via Stata's merge engine, and keeps the mapping in a format that is already part of any Stata workflow.

In practice, switching from string IDs to numeric using `naam id` can reduce file sizes dramatically — on one administrative dataset with multiple ID columns, this brought a 16 GB file down to 5 GB.

---

## Requirements

Stata 14 or higher. No user-written dependencies.

---

## Citation

If you use `naam`, please also cite the package that inspired it:

> Das, Kishor K. (2014). "CODEBOOKOUT: Stata module to save codebook in MS excel format." *Statistical Software Components* S457811, Boston College Department of Economics. https://ideas.repec.org/c/boc/bocode/s457811.html

---

## Author

Vijayshree Jayaraman
[jvijayshree26@gmail.com](mailto:jvijayshree26@gmail.com)
