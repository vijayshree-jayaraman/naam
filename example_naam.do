* =============================================================================
* naam: example workflow (v1.1.1)
* =============================================================================
* This do-file demonstrates the full naam workflow using the two sample
* datasets installed with the package: naam_round1.dta and naam_round2.dta
* (200 households each, with district, occupation, religion, and an
* alphanumeric household ID hhid in the format HH-MH-NNNNN).
*
* Round 2 introduces a new district (Amravati) not present in Round 1,
* which would silently break a naive encode + append.
* =============================================================================

* Install naam if not already installed
* ssc install naam, all
* Development version:
* net install naam, from("https://raw.githubusercontent.com/vijayshree-jayaraman/naam/main/")


* -----------------------------------------------------------------------------
* PART 1: Encode Round 1 and save mappings
* -----------------------------------------------------------------------------

use naam_round1.dta, clear

* Encode string variables and save mappings to Excel
naam encode district occupation religion using naam_maps.xlsx, replace

* Convert string household ID to numeric; saves naam_ids_hhid.dta
naam id hhid using naam_ids, replace keep

* ,keep retains the original string ID as _str_hhid for reference
* the numeric hhid is now in the dataset

save naam_enc1.dta, replace


* -----------------------------------------------------------------------------
* PART 2: Apply same mappings to Round 2
* -----------------------------------------------------------------------------

use naam_round2.dta, clear

* Apply saved mappings -- Amravati is new, gets next available code
naam apply using naam_maps.xlsx

* Convert IDs: known IDs get same codes, new IDs get next available codes
naam id hhid using naam_ids, keep


* -----------------------------------------------------------------------------
* PART 3: Append -- codes are now consistent across both rounds
* -----------------------------------------------------------------------------

append using naam_enc1.dta

* Every district code is consistent -- no silent miscodes
tab district

* Inspect full mapping from inside Stata
naam list using naam_maps.xlsx


* -----------------------------------------------------------------------------
* PART 4: Quality checks
* -----------------------------------------------------------------------------

* Check that in-memory labels match the saved mapping
naam check district occupation religion using naam_maps.xlsx

* Decode a variable back to its original strings (uses Excel mapping)
naam decode district using naam_maps.xlsx
tab district


* -----------------------------------------------------------------------------
* PART 5: ID import -- convert numeric IDs back to original strings
* -----------------------------------------------------------------------------
* After naam id, hhid is stored as a numeric. Use naam id import to get
* the original string IDs back. This is the counterpart to naam id.

* Option A: replace hhid with its original string (default)
naam id import hhid using naam_ids
list hhid in 1/5

* Option B: keep the numeric as _num_hhid and restore string to hhid
* naam id import hhid using naam_ids, keep

* Option C: keep numeric hhid intact; create new variable hhid_str alongside it
* (,suffix and ,keep are mutually exclusive)
* naam id import hhid using naam_ids, suffix(_str)


* -----------------------------------------------------------------------------
* PART 6: Export and restore labels (naam export / naam apply workflow)
* -----------------------------------------------------------------------------

sysuse auto, clear

* Save value labels before they get stripped
naam export using naam_labels.xlsx, replace

* Simulate receiving data without labels
label drop _all
tab foreign    // shows 0 and 1, no labels

* Restore with a single command
naam apply using naam_labels.xlsx
tab foreign    // labels restored


* -----------------------------------------------------------------------------
* PART 7: Compare two mapping files (QA across rounds)
* -----------------------------------------------------------------------------

* After processing multiple rounds, compare their mappings for consistency
* naam compare using naam_maps_round1.xlsx, using2(naam_maps_round2.xlsx)
