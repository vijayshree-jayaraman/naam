*! naam version 1.1.1
*! Author: Vijayshree Jayaraman (jvijayshree26@gmail.com)
*! GitHub: https://github.com/vijayshree-jayaraman
*! "What's in a naam?" -- consistent encoding, ID hashing, label management
*!
*! Inspired by codebookout (Das, 2014):
*!   Das, Kishor K. (2014). "CODEBOOKOUT: Stata module to save codebook in
*!   MS excel format." Statistical Software Components S457811,
*!   Boston College Department of Economics.
*!   RePEC: boc:bocode:s457811
*!   https://ideas.repec.org/c/boc/bocode/s457811.html
*!
*! Subcommands:
*!   naam encode  : encode string vars and save exact mappings to Excel
*!                   (,all option also saves numeric vars with value labels)
*!   naam apply   : reapply saved mappings to any file instantly
*!   naam id      : convert alphanumeric IDs to consistent numerics
*!   naam id import: convert numeric IDs back to original strings (new)
*!   naam export  : save labels from already-encoded datasets
*!   naam list    : inspect a mapping file from inside Stata
*!   naam decode  : reverse encoding -- numeric back to string
*!   naam check   : compare in-memory labels against saved mapping
*!   naam compare : compare two mapping files against each other

program define naam
    version 14.0
    local subcmd = word("`0'", 1)
    local rest   = substr(`"`0'"', length("`subcmd'") + 2, .)

    * Check for "id import" as a two-word subcommand
    if "`subcmd'" == "id" {
        local sub2 = word("`rest'", 1)
        if "`sub2'" == "import" {
            local rest2 = substr(`"`rest'"', length("`sub2'") + 2, .)
            naam_id_import `rest2'
            exit
        }
    }

    if "`subcmd'" == "encode" {
        naam_encode `rest'
    }
    else if "`subcmd'" == "id" {
        naam_id `rest'
    }
    else if "`subcmd'" == "export" {
        naam_export `rest'
    }
    else if "`subcmd'" == "apply" {
        naam_apply `rest'
    }
    else if "`subcmd'" == "list" {
        naam_list `rest'
    }
    else if "`subcmd'" == "decode" {
        naam_decode `rest'
    }
    else if "`subcmd'" == "check" {
        naam_check `rest'
    }
    else if "`subcmd'" == "compare" {
        naam_compare `rest'
    }
    else {
        di as err "Subcommand must be: encode, apply, id, id import, export, list, decode, check, or compare"
        exit 198
    }
end


* -----------------------------------------------------------------------------
program define naam_encode
* Encode string variables to numeric and save exact mappings to Excel.
* When ,all is specified, also saves numeric variables with value labels
* as type=export in the same xlsx -- no need for a separate naam export call.
* -----------------------------------------------------------------------------
    version 14.0
    syntax varlist using/ [, replace keep ALL]
    local fname `"`using'"'
    if substr(`"`fname'"',-5,5)!=".xlsx" & substr(`"`fname'"',-4,4)!=".xls" {
        local fname `"`fname'.xlsx"'
    }
    local nvars : word count `varlist'
    tempvar _naam_order
    quietly gen long `_naam_order' = _n
    local i 1
    foreach v of local varlist {
        local vtype : type `v'
        if substr("`vtype'",1,3) != "str" {
            di as txt "  (skipping `v': not a string)"
            local m_name_`i' "`v'"
            local m_nvals_`i' 0
            local ++i
            continue
        }
        local m_name_`i' "`v'"
        local m_vl_`i' : variable label `v'
        local m_type_`i' : type `v'

        * Check unique value count before levelsof.
        * Stata's encode is limited to 65,536 categories, and levelsof
        * will exhaust macro memory on high-cardinality variables like IDs.
        quietly {
            tempvar _nuniq_flag
            bysort `v': gen byte `_nuniq_flag' = (_n == 1)
            count if `_nuniq_flag' == 1
            local nuniq = r(N)
            drop `_nuniq_flag'
        }
        if `nuniq' > 65536 {
            di as err "  `v': `nuniq' unique values -- too many for naam encode (limit: 65,536)."
            di as err "  If `v' is an ID variable use: naam id `v' using filename, replace"
            local m_name_`i' "`v'"
            local m_nvals_`i' 0
            local ++i
            continue
        }
        quietly levelsof `v', local(vals)
        local m_nvals_`i' 0
        local code 1
        foreach val of local vals {
            local ++m_nvals_`i'
            local m_code_`i'_`m_nvals_`i'' `code'
            local m_val_`i'_`m_nvals_`i'' `"`val'"'
            local ++code
        }
        local ++i
    }
    quietly sort `_naam_order'

    * Preflight Excel sheet names before mutating the data.
    preserve
    quietly {
        clear
        set obs `nvars'
        gen str32 varname = ""
        gen str31 sheetname = ""
        forval i = 1/`nvars' {
            replace varname = "`m_name_`i''" in `i'
            replace sheetname = substr("`m_name_`i''",1,31) in `i'
        }
        drop if varname == ""
        duplicates tag sheetname, gen(_naam_sheetdup)
        count if _naam_sheetdup > 0
        if r(N) > 0 {
            noi di as err "Excel sheet-name collision detected after 31-character truncation."
            noi di as err "Rename one of the colliding variables before using naam encode."
            restore
            quietly drop `_naam_order'
            exit 459
        }
    }
    restore

    * Preflight against an existing workbook too; do this before encoding.
    capture confirm file `"`fname'"'
    if !_rc {
        preserve
        capture quietly {
            import excel using `"`fname'"', sheet("index") firstrow clear allstring
            confirm variable varname
            capture confirm variable sheetname
            if _rc {
                gen str31 sheetname = substr(varname,1,31)
            }
            keep varname sheetname
            forval i = 1/`nvars' {
                drop if varname == "`m_name_`i''"
            }
            local oldN = _N
            local newN = `oldN' + `nvars'
            set obs `newN'
            forval i = 1/`nvars' {
                local row = `oldN' + `i'
                replace varname = "`m_name_`i''" in `row'
                replace sheetname = substr("`m_name_`i''",1,31) in `row'
            }
            duplicates tag sheetname, gen(_naam_sheetdup)
            count if _naam_sheetdup > 0
            if r(N) > 0 error 459
        }
        local preflight_rc = _rc
        restore
        if `preflight_rc' {
            di as err "Excel sheet-name collision or malformed index detected before encoding."
            di as err "Rename one of the colliding variables or fix the workbook index."
            quietly drop `_naam_order'
            exit 459
        }
    }

    local i 1
    foreach v of local varlist {
        local vtype : type `v'
        if substr("`vtype'",1,3) != "str" {
            local ++i
            continue
        }
        if "`keep'" != "" {
            local keepname "_str_`v'"
            if length("`keepname'") > 32 {
                local vstub = substr("`v'",1,24)
                local vsuf : display %02.0f `i'
                local keepname "_str_`vstub'_`vsuf'"
            }
            capture confirm variable `keepname'
            if !_rc {
                di as err "  Cannot create backup: variable `keepname' already exists. Rename it first."
                local m_nvals_`i' 0
                local ++i
                continue
            }
            quietly clonevar `keepname' = `v'
            label var `keepname' "Original string: `v'"
        }
        * Use a prefixed label name to avoid colliding with any existing
        * label named `v' that may already be defined in memory.
        * Drop the _naam_ label first if it exists so encode starts clean.
        tempvar enc_tmp
        tempname enc_label
        capture label drop `enc_label'
        quietly encode `v', gen(`enc_tmp') label(`enc_label')
        quietly drop `v'
        quietly rename `enc_tmp' `v'
        local ++i
    }

    * --- If ,all specified: collect numeric vars with value labels -----------
    local n_export 0
    if "`all'" != "" {
        quietly ds
        foreach v of varlist `r(varlist)' {
            * Skip vars already in the encode varlist
            local already 0
            forval i = 1/`nvars' {
                if "`m_name_`i''" == "`v'" local already 1
            }
            if `already' continue
            * Skip strings
            local vtype : type `v'
            if substr("`vtype'",1,3) == "str" continue
            * Skip ID vars (no value label, high cardinality)
            local lbname : value label `v'
            if "`lbname'" == "" continue
            * Collect label mapping
            local ++n_export
            local ex_name_`n_export' "`v'"
            local ex_vl_`n_export'   : variable label `v'
            local ex_type_`n_export' : type `v'
            local ex_lbn_`n_export'  "`lbname'"
            quietly label list `lbname'
            local kmin = r(min)
            local kmax = r(max)
            local ex_nvals_`n_export' 0
            forval code = `kmin'/`kmax' {
                local txt : label `lbname' `code', strict
                if `"`txt'"' != "" {
                    local ++ex_nvals_`n_export'
                    local ex_code_`n_export'_`ex_nvals_`n_export'' `code'
                    local ex_txt_`n_export'_`ex_nvals_`n_export''  `"`txt'"'
                }
            }
        }
        if `n_export' > 0 {
            di as txt "  ,all: found `n_export' numeric variable(s) with value labels to export."
        }
    }

    preserve
    quietly {
        clear
        local nvalid 0
        forval i = 1/`nvars' {
            if `m_nvals_`i'' > 0 {
                local ++nvalid
            }
        }
        local ntotal = `nvalid' + `n_export'
        if `ntotal' == 0 {
            restore
            quietly drop `_naam_order'
            di as err "No string variables found"
            exit 109
        }
        set obs `ntotal'
        gen str32  varname  = ""
        gen str244 varlabel = ""
        gen str16  vartype  = ""
        gen str32  lblname  = ""
        gen str10  type     = ""
        gen str31  sheetname = ""
        * Write encode entries
        local j 1
        forval i = 1/`nvars' {
            if `m_nvals_`i'' == 0 {
                continue
            }
            replace varname  = "`m_name_`i''"  in `j'
            replace varlabel = `"`m_vl_`i''"'  in `j'
            replace vartype  = `"`m_type_`i''"' in `j'
            replace type     = "encode"         in `j'
            replace sheetname = substr("`m_name_`i''",1,31) in `j'
            local ++j
        }
        * Write export entries
        forval i = 1/`n_export' {
            replace varname  = "`ex_name_`i''"  in `j'
            replace varlabel = `"`ex_vl_`i''"'  in `j'
            replace vartype  = `"`ex_type_`i''"' in `j'
            replace lblname  = `"`ex_lbn_`i''"' in `j'
            replace type     = "export"          in `j'
            replace sheetname = substr("`ex_name_`i''",1,31) in `j'
            local ++j
        }
        * Merge with existing index so prior variables are not wiped
        tempfile newidx
        save `"`newidx'"', replace
        capture confirm file `"`fname'"'
        if !_rc {
            capture {
                import excel using `"`fname'"', sheet("index") firstrow clear allstring
                confirm variable varname
                capture confirm variable sheetname
                if _rc {
                    gen str31 sheetname = substr(varname,1,31)
                }
                capture confirm variable vartype
                if _rc {
                    gen str16 vartype = ""
                }
                capture confirm variable lblname
                if _rc {
                    gen str32 lblname = ""
                }
                * Drop rows for variables being updated in this call
                forval i = 1/`nvars' {
                    if `m_nvals_`i'' > 0 {
                        drop if varname == "`m_name_`i''"
                    }
                }
                forval i = 1/`n_export' {
                    drop if varname == "`ex_name_`i''"
                }
                append using `"`newidx'"'
            }
            if _rc {
                use `"`newidx'"', clear
            }
        }
        else {
            use `"`newidx'"', clear
        }
        duplicates tag sheetname, gen(_naam_sheetdup)
        count if _naam_sheetdup > 0
        if r(N) > 0 {
            noi di as err "Excel sheet-name collision detected after 31-character truncation."
            noi di as err "Rename one of the colliding variables before using naam encode/export."
            restore
            quietly drop `_naam_order'
            exit 459
        }
        drop _naam_sheetdup
        export excel varname varlabel vartype lblname type sheetname using `"`fname'"', ///
            sheet("index") sheetreplace firstrow(variables)
    }
    restore
    di as txt "  -> [index] written."

    * Write encode mapping sheets
    forval i = 1/`nvars' {
        if `m_nvals_`i'' == 0 {
            continue
        }
        preserve
        quietly {
            clear
            set obs `m_nvals_`i''
            gen long   numeric_code = .
            gen strL   string_value = ""
            forval e = 1/`m_nvals_`i'' {
                * BUG FIX: store as numeric, not string, so merges work correctly
                replace numeric_code = `m_code_`i'_`e''  in `e'
                replace string_value = `"`m_val_`i'_`e''"' in `e'
            }
            local shname = substr("`m_name_`i''",1,31)
            export excel numeric_code string_value using `"`fname'"', ///
                sheet("`shname'") sheetreplace firstrow(variables)
        }
        restore
        di as txt "  -> [`m_name_`i''] written (`m_nvals_`i'' categories)."
    }

    * Write export label sheets
    forval i = 1/`n_export' {
        if `ex_nvals_`i'' == 0 continue
        preserve
        quietly {
            clear
            set obs `ex_nvals_`i''
            gen long   numeric_code = .
            gen strL   string_value = ""
            forval e = 1/`ex_nvals_`i'' {
                replace numeric_code = `ex_code_`i'_`e''   in `e'
                replace string_value = `"`ex_txt_`i'_`e''"'  in `e'
            }
            local shname = substr("`ex_name_`i''",1,31)
            export excel numeric_code string_value using `"`fname'"', ///
                sheet("`shname'") sheetreplace firstrow(variables)
        }
        restore
        di as txt "  -> [`ex_name_`i''] exported (`ex_nvals_`i'' labels)."
    }

    quietly sort `_naam_order'
    quietly drop `_naam_order'
    di as res `"naam encode complete -> `fname'"'
end


* -----------------------------------------------------------------------------
program define naam_id
* Convert a string ID variable to consistent numeric codes across files.
* On first call: assigns codes 1, 2, 3... alphabetically and saves mapping.
* On subsequent calls: looks up saved mapping, adds new codes for new IDs.
* Accepts a varlist: each variable is processed in turn, each saved to its
* own .dta file (base_varname.dta). No row-count limit.
* -----------------------------------------------------------------------------
    version 14.0
    syntax varlist using/ [, replace keep strict]

    * -- Strip any extension to get a clean base name -------------------------
    local base `"`using'"'
    if substr(`"`base'"', -4, 4) == ".dta" {
        local base = substr(`"`base'"', 1, length(`"`base'"') - 4)
    }
    else if substr(`"`base'"', -5, 5) == ".xlsx" {
        local base = substr(`"`base'"', 1, length(`"`base'"') - 5)
    }
    else if substr(`"`base'"', -4, 4) == ".xls" {
        local base = substr(`"`base'"', 1, length(`"`base'"') - 4)
    }

    * -- Process each variable in the varlist ---------------------------------
    local n_processed 0  // counts variables actually completed
    local vidx 0
    foreach v of local varlist {
        local ++vidx

        local vtype : type `v'
        if substr("`vtype'",1,3) != "str" {
            di as err "  `v' is not a string variable -- skipping"
            di as err "  naam id requires a string variable. For ID conversion use a string ID."
            continue
        }

        * One .dta per variable: base_varname.dta
        local fname `"`base'_`v'.dta"'

        * -- Check whether an output mapping already exists --------------------
        * A valid existing mapping is updated in place. ,replace is only needed
        * when the existing mapping is invalid and the user wants to start over.
        capture confirm file `"`fname'"'
        local file_rc = _rc

        local vl : variable label `v'
        quietly count
        di as txt "  `r(N)' observations, processing `v'..."

        * -- Check for intermediate variable name collisions before touching dataset --
        * string_value and numeric_code are reserved during the existing-map merge.
        local collision 0
        foreach chkvar in string_value numeric_code {
            capture confirm variable `chkvar'
            if !_rc {
                di as err "  Cannot process `v': variable `chkvar' already exists. Rename it first."
                local collision 1
            }
        }
        local keepname ""
        if "`keep'" != "" {
            local keepname "_str_`v'"
            if length("`keepname'") > 32 {
                local vstub = substr("`v'",1,24)
                local vsuf : display %02.0f `vidx'
                local keepname "_str_`vstub'_`vsuf'"
            }
            capture confirm variable `keepname'
            if !_rc {
                di as err "  Cannot keep original string: variable `keepname' already exists. Rename it first."
                local collision 1
            }
        }
        if `collision' continue

        * Always keep a string backup for mapping and for ,keep option.
        tempvar str_tmp id_tmp _naam_order
        quietly gen long `_naam_order' = _n
        quietly clonevar `str_tmp' = `v'

        * -- Warn and note missing (blank) string IDs -------------------------
        quietly count if missing(`v')
        local nmiss = r(N)
        if `nmiss' > 0 {
            di as txt "  warning: `nmiss' observation(s) have a missing (blank) ID in `v'."
            di as txt "  These will remain missing (.) in the numeric variable."
        }

        * -- Check if a valid mapping .dta already exists for this variable ---
        * Normalize duplicate string_value rows before m:1 merge. Refuse any
        * mapping where one numeric_code maps to multiple strings.
        tempfile existing_map
        local file_exists 0
        local max_existing 0
        if !`file_rc' {
            preserve
            capture {
                quietly use `"`fname'"', clear
                confirm variable string_value
                confirm variable numeric_code
                capture confirm numeric variable numeric_code
                if _rc {
                    destring numeric_code, replace
                }
                keep string_value numeric_code
                drop if missing(string_value)
                drop if missing(numeric_code)
                duplicates drop string_value numeric_code, force
                isid string_value
                isid numeric_code
                quietly count
                if r(N) == 0 error 1
                quietly summarize numeric_code, meanonly
                local max_existing = r(max)
                sort numeric_code
                save `"`existing_map'"', replace
            }
            if !_rc local file_exists 1
            restore
            if !`file_exists' & "`replace'" == "" {
                di as err "  Mapping file `fname' exists but is invalid or internally inconsistent."
                di as err "  Fix it manually, or specify ,replace to discard it and create a fresh mapping."
                quietly drop `str_tmp' `_naam_order'
                continue
            }
        }

        * -- CASE 1: No valid mapping -- assign fresh codes -------------------
        if !`file_exists' {
            di as txt "  No valid mapping found -- assigning fresh codes."
            if "`strict'" != "" {
                di as err "strict: no valid saved mapping found for `v'."
                di as err "Remove ,strict to create a fresh mapping."
                quietly drop `str_tmp' `_naam_order'
                exit 459
            }
            quietly {
                * BUG FIX: sort before group to ensure alphabetical, stable assignment
                * Exclude missing (blank) string values from group() -- they stay missing.
                tempvar grp
                sort `v'
                egen `grp' = group(`v') if !missing(`v')
                gen double `id_tmp' = `grp'
                drop `grp'
            }
        }

        * -- CASE 2: Valid mapping exists -- look up, add new if needed -------
        else {
            di as txt "  Existing mapping found -- looking up codes..."

            capture quietly {
                * BUG FIX: rename to string_value (match mapping key name)
                * then merge, then rename back -- prevents variable name collision
                rename `v' string_value
                merge m:1 string_value using `"`existing_map'"', ///
                    keepusing(numeric_code) nogen keep(1 3)
                rename string_value `v'
            }
            if _rc {
                local merge_rc = _rc
                capture rename string_value `v'
                capture drop numeric_code
                quietly sort `_naam_order'
                quietly drop `str_tmp' `_naam_order'
                di as err "  Merge failed while assigning IDs for `v' (r(`merge_rc')). Dataset restored for this variable."
                continue
            }

            quietly count if missing(numeric_code) & !missing(`v')
            local nnew = r(N)

            if `nnew' > 0 & "`strict'" != "" {
                di as err "strict: `nnew' new IDs not in saved mapping for `v'."
                di as err "Remove ,strict to allow new IDs to be added."
                * Clean up variables added during this variable's processing
                * so the dataset is returned to its pre-call state for `v'.
                * Note: previously processed variables in the loop are already
                * encoded -- strict applies per-variable, not to the whole call.
                quietly sort `_naam_order'
                quietly drop numeric_code `str_tmp' `_naam_order'
                exit 459
            }

            if `nnew' > 0 {
                di as txt "  `nnew' new IDs -- assigning new codes."
                quietly {
                    tempvar newgrp
                    sort `v'
                    * Exclude missing string values -- they stay missing in numeric_code.
                    egen `newgrp' = group(`v') if missing(numeric_code) & !missing(`v')
                    replace numeric_code = `max_existing' + `newgrp' ///
                        if missing(numeric_code) & !missing(`v')
                    drop `newgrp'
                }
            }
            else {
                di as txt "  All IDs matched. No new codes needed."
            }

            quietly gen double `id_tmp' = numeric_code
            quietly drop numeric_code
        }

        * -- Build updated mapping ---------------------------------------------
        * Build from current data plus the full prior mapping so IDs absent from
        * the current file are retained and numeric codes are never reused.
        tempfile new_mapping
        preserve
        quietly {
            keep `str_tmp' `id_tmp'
            rename `str_tmp' string_value
            rename `id_tmp'  numeric_code
            * The dataset may have many rows per unique ID. Collapse to one
            * row per unique string_value for the mapping file.
            * numeric_code uniqueness is guaranteed by construction: fresh
            * codes come from egen group() (unique by definition), and new
            * codes on subsequent calls are maxcode + newgrp (also unique).
            * Exclude blank/missing string IDs -- they have no valid code.
            drop if missing(string_value)
            duplicates drop string_value, force
            if `file_exists' {
                append using `"`existing_map'"'
                duplicates drop string_value, force
            }
            isid string_value
            isid numeric_code
            sort numeric_code
            local ntotal = _N
            save `"`new_mapping'"', replace
        }
        restore

        * -- Apply to dataset --------------------------------------------------
        quietly drop `v'
        quietly rename `id_tmp' `v'
        quietly recast double `v'
        label var `v' "`vl'"

        if "`keep'" == "" {
            quietly drop `str_tmp'
        }
        else {
            quietly rename `str_tmp' `keepname'
            label var `keepname' "Original string ID: `v'"
        }
        quietly sort `_naam_order'
        quietly drop `_naam_order'

        di as txt "  -> `v' assigned: `ntotal' unique IDs in mapping."

        * -- Save mapping as .dta (no row-count limit) ------------------------
        preserve
        quietly {
            use `"`new_mapping'"', clear
            save `"`fname'"', replace
        }
        restore
        local ++n_processed
        di as res `"naam id complete for `v' -> `fname'"'
    }

    if `n_processed' == 0 {
        di as err "No variables were successfully processed."
        di as err "Check that variables are strings and that any existing mapping files are valid."
        exit 109
    }
end


* -----------------------------------------------------------------------------
program define naam_id_import
* NEW SUBCOMMAND: naam id import
* Convert numeric ID variable(s) back to their original strings using the
* saved .dta mapping files (base_varname.dta).
*
* Default behaviour (no options):
*   varname is replaced in-place with its original string values.
*
* , keep
*   The numeric variable is retained as _num_varname; the string replaces
*   varname in-place.
*
* , suffix(str)
*   Instead of replacing varname in-place, the string is stored in a new
*   variable named varname+suffix (e.g. hhid_str). The original numeric
*   varname is always kept. Cannot be combined with ,keep.
*
* Syntax: naam id import varlist using basepath [, keep suffix(str)]
* -----------------------------------------------------------------------------
    version 14.0
    syntax varlist using/ [, keep SUFfix(string)]

    * suffix and keep are mutually exclusive
    if "`keep'" != "" & "`suffix'" != "" {
        di as err "Options ,keep and ,suffix() are mutually exclusive."
        di as err "  ,keep   : replace varname with string; keep numeric as _num_varname"
        di as err "  ,suffix : keep numeric varname; create new string variable varname+suffix"
        exit 198
    }

    * Strip any extension from base path (consistent with naam id)
    local base `"`using'"'
    if substr(`"`base'"', -4, 4) == ".dta" {
        local base = substr(`"`base'"', 1, length(`"`base'"') - 4)
    }
    else if substr(`"`base'"', -5, 5) == ".xlsx" {
        local base = substr(`"`base'"', 1, length(`"`base'"') - 5)
    }
    else if substr(`"`base'"', -4, 4) == ".xls" {
        local base = substr(`"`base'"', 1, length(`"`base'"') - 4)
    }

    local n_done 0
    foreach v of local varlist {

        * Must be numeric
        local vtype : type `v'
        if substr("`vtype'",1,3) == "str" {
            di as txt "  (skipping `v': already a string -- use naam decode for encode mappings)"
            continue
        }

        local fname `"`base'_`v'.dta"'
        capture confirm file `"`fname'"'
        if _rc {
            di as err "  Mapping file not found for `v': `fname' -- skipping"
            continue
        }

        * Validate and normalize mapping file before touching the dataset.
        tempfile clean_id_map
        preserve
        capture {
            quietly use `"`fname'"', clear
            confirm variable string_value
            confirm variable numeric_code
            capture confirm numeric variable numeric_code
            if _rc {
                destring numeric_code, replace
            }
            keep string_value numeric_code
            drop if missing(string_value)
            drop if missing(numeric_code)
            duplicates drop string_value numeric_code, force
            isid string_value
            isid numeric_code
            save `"`clean_id_map'"', replace
        }
        if _rc {
            restore
            di as err "  Mapping file `fname' is invalid or internally inconsistent -- skipping"
            continue
        }
        restore

        local vl : variable label `v'
        tempvar _naam_order
        quietly gen long `_naam_order' = _n

        * Determine output variable name and check for collision
        if "`suffix'" != "" {
            * ,suffix mode: numeric varname stays; new string var is varname+suffix
            local strvar "`v'`suffix'"
            capture confirm new variable `strvar'
            if _rc {
                di as err "  Cannot create variable `strvar'. Choose a shorter/different suffix()."
                quietly drop `_naam_order'
                continue
            }
        }
        else {
            * default / ,keep mode: string will go into varname; use a temp name during merge
            tempvar strvar
        }

        local numkeep ""
        if "`keep'" != "" {
            local numkeep "_num_`v'"
            if length("`numkeep'") > 32 {
                local vstub = substr("`v'",1,24)
                local numkeep "_num_`vstub'_01"
            }
            capture confirm variable `numkeep'
            if !_rc {
                di as err "  Cannot keep numeric: variable `numkeep' already exists. Rename it first."
                quietly drop `_naam_order'
                continue
            }
        }

        * Merge string values in using numeric_code as key.
        * Check that neither 'numeric_code' nor 'string_value' already exist.
        local merge_collision 0
        foreach chkvar in numeric_code string_value {
            capture confirm variable `chkvar'
            if !_rc {
                di as err "  Cannot process `v': variable `chkvar' already exists in dataset."
                di as err "  Rename or drop it before running naam id import."
                local merge_collision 1
            }
        }
        if `merge_collision' {
            quietly drop `_naam_order'
            continue
        }
        capture quietly {
            rename `v' numeric_code
            merge m:1 numeric_code using `"`clean_id_map'"', ///
                keepusing(string_value) nogen keep(1 3)
            rename string_value `strvar'
            rename numeric_code `v'
        }
        if _rc {
            local merge_rc = _rc
            capture rename numeric_code `v'
            capture drop string_value
            capture drop `strvar'
            di as err "  Merge failed while importing IDs for `v' (r(`merge_rc')). Dataset restored for this variable."
            quietly sort `_naam_order'
            quietly drop `_naam_order'
            continue
        }

        quietly count if `v' != . & `strvar' == ""
        local nunmatched = r(N)
        if `nunmatched' > 0 {
            di as txt "  warning: `nunmatched' observation(s) in `v' had no match in the mapping."
            if "`suffix'" == "" {
                di as err "  Refusing to replace `v' because unmatched numeric IDs would become blank strings."
                di as err "  Use ,suffix() to inspect unmatched IDs while keeping the numeric variable."
                quietly drop `strvar'
                quietly sort `_naam_order'
                quietly drop `_naam_order'
                exit 459
            }
        }

        if "`suffix'" != "" {
            * ,suffix mode: numeric stays as-is; label the new string var
            label var `strvar' "`vl' (string)"
            quietly sort `_naam_order'
            quietly drop `_naam_order'
            di as txt "  -> `strvar' created as string; `v' kept as numeric."
        }
        else if "`keep'" != "" {
            * ,keep mode: rename numeric to _num_v; rename string to v
            quietly rename `v' `numkeep'
            quietly rename `strvar' `v'
            label var `v' "`vl'"
            label var `numkeep' "Numeric ID: `v'"
            quietly sort `_naam_order'
            quietly drop `_naam_order'
            di as txt "  -> `v' converted to string; numeric kept as `numkeep'."
        }
        else {
            * default mode: drop numeric; rename string to v
            quietly drop `v'
            quietly rename `strvar' `v'
            label var `v' "`vl'"
            quietly sort `_naam_order'
            quietly drop `_naam_order'
            di as txt "  -> `v' converted back to string."
        }

        local ++n_done
    }

    if `n_done' == 0 {
        di as err "No variables were successfully imported."
        exit 109
    }
    di as res "naam id import complete. `n_done' variable(s) converted."
end


* -----------------------------------------------------------------------------
program define naam_export
* Save variable labels and value labels from an already-encoded dataset
* to Excel so they can be reattached later with naam apply.
* -----------------------------------------------------------------------------
    version 14.0
    syntax using/ [, replace]
    local fname `"`using'"'
    if substr(`"`fname'"',-5,5)!=".xlsx" & substr(`"`fname'"',-4,4)!=".xls" {
        local fname `"`fname'.xlsx"'
    }
    quietly ds
    local allvars `r(varlist)'
    local nvars : word count `allvars'
    local i 1
    foreach v of local allvars {
        local m_name_`i' "`v'"
        local m_lbl_`i'  : variable label `v'
        local m_type_`i' : type `v'
        local m_lbn_`i'  : value label `v'
        local lbname `"`m_lbn_`i''"'
        local m_nvals_`i' 0
        if `"`lbname'"' != "" {
            quietly label list `lbname'
            local kmin = r(min)
            local kmax = r(max)
            local entry 0
            forval code = `kmin'/`kmax' {
                local txt : label `lbname' `code', strict
                if `"`txt'"' != "" {
                    local ++entry
                    local m_vcode_`i'_`entry' `code'
                    local m_vtxt_`i'_`entry'  `"`txt'"'
                }
            }
            local m_nvals_`i' `entry'
        }
        local ++i
    }
    preserve
    quietly {
        clear
        set obs `nvars'
        gen str32  varname  = ""
        gen str244 varlabel = ""
        gen str16  vartype  = ""
        gen str32  lblname  = ""
        gen str10  type     = "export"
        gen str31  sheetname = ""
        forval i = 1/`nvars' {
            replace varname  = `"`m_name_`i''"'  in `i'
            replace varlabel = `"`m_lbl_`i''"'   in `i'
            replace vartype  = `"`m_type_`i''"'  in `i'
            replace lblname  = `"`m_lbn_`i''"'   in `i'
            replace sheetname = substr(`"`m_name_`i''"',1,31) in `i'
        }
        duplicates tag sheetname, gen(_naam_sheetdup)
        count if _naam_sheetdup > 0
        if r(N) > 0 {
            noi di as err "Excel sheet-name collision detected after 31-character truncation."
            noi di as err "Rename one of the colliding variables before using naam export."
            restore
            exit 459
        }
        drop _naam_sheetdup
        export excel varname varlabel vartype lblname type sheetname using `"`fname'"', ///
            sheet("index") sheetreplace firstrow(variables)
    }
    restore
    di as txt "  -> [index] written."
    local nsheets 0
    forval i = 1/`nvars' {
        if `m_nvals_`i'' == 0 {
            continue
        }
        preserve
        quietly {
            clear
            set obs `m_nvals_`i''
            gen long   numeric_code = .
            gen strL   string_value = ""
            forval e = 1/`m_nvals_`i'' {
                replace numeric_code = `m_vcode_`i'_`e''   in `e'
                replace string_value = `"`m_vtxt_`i'_`e''"'  in `e'
            }
            local shname = substr(`"`m_name_`i''"',1,31)
            export excel numeric_code string_value using `"`fname'"', ///
                sheet("`shname'") sheetreplace firstrow(variables)
        }
        restore
        local ++nsheets
    }
    if `nsheets' == 0 {
        di as txt "  (no value-labeled variables found)"
    }
    di as res `"naam export complete -> `fname'"'
end


* -----------------------------------------------------------------------------
program define naam_apply
* Read a naam Excel file and reattach all saved mappings to the dataset
* currently in memory. Handles type=encode, type=export, and type=id.
* -----------------------------------------------------------------------------
    version 14.0
    syntax using/ [, VARSOnly LABELSOnly]
    if "`varsonly'" != "" & "`labelsonly'" != "" {
        di as err "Options ,varsonly and ,labelsonly are mutually exclusive."
        exit 198
    }
    local fname `"`using'"'
    if substr(`"`fname'"',-5,5)!=".xlsx" & substr(`"`fname'"',-4,4)!=".xls" {
        local fname `"`fname'.xlsx"'
    }
    confirm file `"`fname'"'

    * Read index into locals
    tempfile idxtmp
    preserve
    quietly {
        import excel using `"`fname'"', sheet("index") firstrow clear allstring
        capture confirm variable varname
        if _rc {
            restore
            di as err "Sheet [index] missing or malformed"
            exit 111
        }
        capture confirm variable sheetname
        if _rc {
            gen str31 sheetname = substr(varname,1,31)
        }
        duplicates tag varname, gen(_naam_vdup)
        duplicates tag sheetname, gen(_naam_sdup)
        count if _naam_vdup > 0 | _naam_sdup > 0
        if r(N) > 0 {
            restore
            di as err "Sheet [index] has duplicate varname or sheetname entries"
            exit 459
        }
        drop _naam_vdup _naam_sdup
        save `"`idxtmp'"', replace
    }
    restore

    preserve
    quietly use `"`idxtmp'"', clear
    capture confirm variable sheetname
    local has_sheetname = !_rc
    local nrows = _N
    forval i = 1/`nrows' {
        local r_vname_`i'   = varname[`i']
        local r_vlabel_`i'  = varlabel[`i']
        local r_type_`i'    = type[`i']
        if `has_sheetname' {
            local r_sheetname_`i' = sheetname[`i']
        }
        else {
            local r_sheetname_`i' = substr("`r_vname_`i''",1,31)
        }
        capture local r_lbname_`i' = lblname[`i']
        if _rc {
            local r_lbname_`i' ""
        }
    }
    restore

    * Reattach variable labels (all types)
    if "`labelsonly'" == "" {
        local n_vl 0
        forval i = 1/`nrows' {
            local v   "`r_vname_`i''"
            local lbl `"`r_vlabel_`i''"'
            capture confirm variable `v'
            if _rc {
                continue
            }
            if `"`lbl'"' == "" {
                continue
            }
            label variable `v' `"`lbl'"'
            local ++n_vl
        }
        di as txt "Variable labels reattached: `n_vl'"
    }

    * Reattach value labels / mappings (type-aware)
    if "`varsonly'" == "" {
        local n_encode 0
        local n_export 0
        local n_id     0

        forval i = 1/`nrows' {
            local v      "`r_vname_`i''"
            local lbname "`r_lbname_`i''"
            local vtype  "`r_type_`i''"
            local shname "`r_sheetname_`i''"

            capture confirm variable `v'
            if _rc {
                continue
            }

            * Read the mapping sheet for this variable
            tempfile valtmp
            local sheet_ok 1
            preserve
            quietly {
                capture {
                    import excel using `"`fname'"', ///
                        sheet("`shname'") firstrow clear allstring
                    confirm variable numeric_code
                    confirm variable string_value
                    save `"`valtmp'"', replace
                }
                if _rc {
                    local sheet_ok 0
                }
            }
            restore
            if !`sheet_ok' {
                continue
            }

            * Load mapping into locals
            preserve
            quietly use `"`valtmp'"', clear
            capture confirm numeric variable numeric_code
            if _rc {
                destring numeric_code, replace
            }
            capture {
                assert !missing(numeric_code)
                isid numeric_code
                if "`vtype'" == "encode" {
                    assert !missing(string_value)
                    isid string_value
                }
            }
            if _rc {
                restore
                di as err "  [`v'] mapping sheet is invalid or internally inconsistent -- skipped"
                continue
            }
            local nval = _N
            forval j = 1/`nval' {
                local c_`j' = numeric_code[`j']
                local l_`j' = string_value[`j']
            }
            restore

            * -- TYPE: encode -------------------------------------------------
            if "`vtype'" == "encode" {
                if "`lbname'" == "" {
                    local lbname "`v'"
                }
                local maxcode 0
                forval j = 1/`nval' {
                    local cj = real("`c_`j''")
                    if `cj' > `maxcode' {
                        local maxcode `cj'
                    }
                }
                * If `v' is still a string (i.e. applying to a raw file before encoding),
                * scan for new categories and encode string-to-numeric using saved mapping.
                * If `v' is already numeric, skip -- just reattach labels below.
                local new_added 0
                local vt : type `v'
                if substr("`vt'",1,3) == "str" {
                    * Scan for values in the data not yet in the saved mapping
                    quietly levelsof `v', local(cur_vals)
                    foreach val of local cur_vals {
                        local found 0
                        forval j = 1/`nval' {
                            if `"`l_`j''"' == `"`val'"' {
                                local found 1
                            }
                        }
                        if !`found' {
                            local ++maxcode
                            local ++nval
                            local c_`nval' `maxcode'
                            local l_`nval' `"`val'"'
                            local ++new_added
                            di as txt "  [`v'] new category added: `val' -> `maxcode'"
                        }
                    }
                    * Encode string to numeric using saved mapping
                    local vl_save : variable label `v'
                    tempvar naam_enc_tmp
                    quietly gen long `naam_enc_tmp' = .
                    forval j = 1/`nval' {
                        local cj = real("`c_`j''")
                        quietly replace `naam_enc_tmp' = `cj' if `v' == `"`l_`j''"'
                    }
                    quietly drop `v'
                    quietly rename `naam_enc_tmp' `v'
                    if `"`vl_save'"' != "" {
                        label variable `v' `"`vl_save'"'
                    }
                    di as txt "  [`v'] string encoded to numeric using saved mapping."
                }
                * Guard against nval==0 (empty mapping edge case)
                if `nval' == 0 continue
                local code1 = real("`c_1'")
                label define `lbname' `code1' `"`l_1'"', replace
                forval j = 2/`nval' {
                    local cj = real("`c_`j''")
                    label define `lbname' `cj' `"`l_`j''"', add
                }
                label values `v' `lbname'
                if `new_added' > 0 {
                    preserve
                    quietly {
                        clear
                        set obs `nval'
                        gen long   numeric_code = .
                        gen strL   string_value = ""
                        forval j = 1/`nval' {
                            replace numeric_code = real("`c_`j''")  in `j'
                            replace string_value = `"`l_`j''"' in `j'
                        }
                        local shname = substr("`v'",1,31)
                        export excel numeric_code string_value ///
                            using `"`fname'"', ///
                            sheet("`shname'") sheetreplace firstrow(variables)
                    }
                    restore
                    di as txt "  [`v'] Excel mapping updated with `new_added' new category(ies)."
                }
                local ++n_encode
            }

            * -- TYPE: export -------------------------------------------------
            else if "`vtype'" == "export" {
                if "`lbname'" == "" {
                    local lbname "`v'"
                }
                * BUG FIX: guard against nval==0
                if `nval' == 0 continue
                local code1 = real("`c_1'")
                label define `lbname' `code1' `"`l_1'"', replace
                forval j = 2/`nval' {
                    local cj = real("`c_`j''")
                    label define `lbname' `cj' `"`l_`j''"', add
                }
                label values `v' `lbname'
                local ++n_export
            }

            * -- TYPE: id -----------------------------------------------------
            else if "`vtype'" == "id" {
                local ++n_id
                di as txt "  [`v'] is an ID variable -- use naam id import to convert back to string."
            }
        }

        if `n_encode' > 0 {
            di as txt "Encode mappings reattached: `n_encode' variable(s)"
        }
        if `n_export' > 0 {
            di as txt "Value labels reattached:    `n_export' variable(s)"
        }
        if `n_id' > 0 {
            di as txt "ID variables noted:         `n_id' variable(s)"
        }
    }

    di as res "naam apply complete."
end


* -----------------------------------------------------------------------------
program define naam_list
* Inspect a naam Excel mapping file from inside Stata.
* Prints a clean summary of every variable and its categories to the
* Results window. No dataset in memory is required or affected.
* -----------------------------------------------------------------------------
    version 14.0
    syntax using/ [, VARiable(string)]
    local fname `"`using'"'
    if substr(`"`fname'"',-5,5)!=".xlsx" & substr(`"`fname'"',-4,4)!=".xls" {
        local fname `"`fname'.xlsx"'
    }
    confirm file `"`fname'"'

    * Read index
    tempfile idxtmp
    preserve
    quietly {
        import excel using `"`fname'"', sheet("index") firstrow clear allstring
        capture confirm variable varname
        if _rc {
            restore
            di as err "Sheet [index] missing or malformed in `fname'"
            exit 111
        }
        save `"`idxtmp'"', replace
    }
    restore

    preserve
    quietly use `"`idxtmp'"', clear
    capture confirm variable sheetname
    local has_sheetname = !_rc
    local nrows = _N
    forval i = 1/`nrows' {
        local r_vname_`i' = varname[`i']
        local r_vlab_`i'  = varlabel[`i']
        local r_type_`i'  = type[`i']
        if `has_sheetname' {
            local r_sheetname_`i' = sheetname[`i']
        }
        else {
            local r_sheetname_`i' = substr("`r_vname_`i''",1,31)
        }
    }
    restore

    * Header
    di as txt _newline "{hline 60}"
    di as res "  naam mapping file: `fname'"
    di as txt "{hline 60}"

    local printed 0
    forval i = 1/`nrows' {
        local v    "`r_vname_`i''"
        local vlab "`r_vlab_`i''"
        local vtyp "`r_type_`i''"
        local shname "`r_sheetname_`i''"

        * If user requested a specific variable, skip others
        if "`variable'" != "" & "`v'" != "`variable'" {
            continue
        }

        * Try to load this variable's mapping sheet
        tempfile valtmp
        local sheet_ok 1
        preserve
        quietly {
            capture {
                import excel using `"`fname'"', ///
                    sheet("`shname'") firstrow clear allstring
                confirm variable numeric_code
                confirm variable string_value
                save `"`valtmp'"', replace
            }
            if _rc local sheet_ok 0
        }
        restore

        di as txt _newline "  Variable : " as res "`v'"
        if `"`vlab'"' != "" {
            di as txt "  Label    : `vlab'"
        }
        di as txt "  Type     : `vtyp'"

        if !`sheet_ok' {
            di as txt "  (no mapping sheet found)"
            local ++printed
            continue
        }

        preserve
        quietly use `"`valtmp'"', clear
        local nval = _N
        restore

        if "`vtyp'" == "id" {
            di as txt "  IDs saved: `nval'"
            di as txt "  (use: naam id import `v' using basepath)"
        }
        else {
            di as txt "  Categories (`nval'):"
            preserve
            quietly use `"`valtmp'"', clear
            forval j = 1/`nval' {
                local cj = numeric_code[`j']
                local lj = string_value[`j']
                di as txt "    `cj'  =  `lj'"
            }
            restore
        }
        local ++printed
    }

    if `printed' == 0 & "`variable'" != "" {
        di as err "  Variable '`variable'' not found in `fname'"
        exit 111
    }

    di as txt _newline "{hline 60}"
    di as txt "  Total variables in file: `nrows'"
    di as txt "{hline 60}"
end


* -----------------------------------------------------------------------------
program define naam_decode
* Reverse a naam encoding: convert numeric variables back to their original
* strings using the saved mapping in the Excel file.
* Works even if value labels have been stripped from the dataset.
* -----------------------------------------------------------------------------
    version 14.0
    syntax varlist using/ [, keep]
    local fname `"`using'"'
    if substr(`"`fname'"',-5,5)!=".xlsx" & substr(`"`fname'"',-4,4)!=".xls" {
        local fname `"`fname'.xlsx"'
    }
    confirm file `"`fname'"'

    foreach v of local varlist {

        * Confirm the variable is numeric
        local vtype : type `v'
        if substr("`vtype'",1,3) == "str" {
            di as txt "  (skipping `v': already a string)"
            continue
        }

        * Try to load the mapping sheet for this variable
        tempfile valtmp
        local shname = substr("`v'",1,31)
        preserve
        capture quietly {
            import excel using `"`fname'"', sheet("index") firstrow clear allstring
            confirm variable sheetname
            count
            local nidx = r(N)
            forval si = 1/`nidx' {
                if varname[`si'] == "`v'" {
                    local shname = sheetname[`si']
                }
            }
        }
        restore
        local sheet_ok 1
        preserve
        quietly {
            capture {
                import excel using `"`fname'"', ///
                    sheet("`shname'") firstrow clear allstring
                confirm variable numeric_code
                confirm variable string_value
                save `"`valtmp'"', replace
            }
            if _rc local sheet_ok 0
        }
        restore

        if !`sheet_ok' {
            di as err "  No mapping sheet found for `v' in `fname' -- skipping"
            continue
        }

        * Load mapping into locals
        preserve
        quietly use `"`valtmp'"', clear
        capture confirm numeric variable numeric_code
        if _rc {
            destring numeric_code, replace
        }
        capture {
            assert !missing(numeric_code)
            isid numeric_code
        }
        if _rc {
            restore
            di as err "  Mapping sheet for `v' is invalid or internally inconsistent -- skipping"
            continue
        }
        local nval = _N
        * BUG FIX: guard against empty mapping sheet
        if `nval' == 0 {
            restore
            di as err "  Mapping sheet for `v' is empty -- skipping"
            continue
        }
        forval j = 1/`nval' {
            local c_`j' = numeric_code[`j']
            local l_`j' = string_value[`j']
        }
        restore

        * Determine the longest string value to set type
        local maxlen 1
        forval j = 1/`nval' {
            local slen = length(`"`l_`j''"')
            if `slen' > `maxlen' {
                local maxlen `slen'
            }
        }
        if `maxlen' > 2045 {
            di as err "  Mapping for `v' contains strings longer than Stata's fixed-string limit."
            di as err "  Decode skipped; export original strings from the source data instead."
            continue
        }

        * Optionally keep the numeric variable
        if "`keep'" != "" {
            * Check for name collision before cloning
            local numkeep "_num_`v'"
            if length("`numkeep'") > 32 {
                local vstub = substr("`v'",1,24)
                local numkeep "_num_`vstub'_01"
            }
            capture confirm variable `numkeep'
            if !_rc {
                di as err "  Cannot keep numeric: variable `numkeep' already exists. Rename it first."
                continue
            }
            quietly clonevar `numkeep' = `v'
            label var `numkeep' "Numeric encoding of `v'"
        }

        * Generate string variable using a tempvar to avoid name collision
        tempvar str_tmp
        quietly gen str`maxlen' `str_tmp' = ""
        forval j = 1/`nval' {
            local cj = real("`c_`j''")
            quietly replace `str_tmp' = `"`l_`j''"' if `v' == `cj'
        }

        * Count unmatched observations
        quietly count if `v' != . & `str_tmp' == ""
        local nunmatched = r(N)
        if `nunmatched' > 0 {
            di as txt "  warning: `nunmatched' observation(s) in `v' had no match in the mapping."
        }

        * Replace original variable with decoded string
        local vlab : variable label `v'
        quietly drop `v'
        quietly rename `str_tmp' `v'
        label var `v' "`vlab'"

        di as txt "  -> `v' decoded to string (`nval' categories mapped)."
    }

    di as res "naam decode complete."
end


* -----------------------------------------------------------------------------
program define naam_check
* Compare the value labels currently attached to variables in memory against
* the saved mapping in the Excel file. Reports matches, conflicts, and any
* categories present in one but not the other.
* Does not modify the dataset or the Excel file.
* -----------------------------------------------------------------------------
    version 14.0
    syntax [anything(name=rawvars)] using/
    local fname `"`using'"'
    if substr(`"`fname'"',-5,5)!=".xlsx" & substr(`"`fname'"',-4,4)!=".xls" {
        local fname `"`fname'.xlsx"'
    }
    confirm file `"`fname'"'

    * Build the varlist -- skip any names that don't exist in the dataset
    local varlist ""
    if "`rawvars'" != "" {
        foreach v of local rawvars {
            capture confirm variable `v'
            if _rc {
                di as txt "  `v': not found in dataset -- skipped"
            }
            else {
                local varlist `varlist' `v'
            }
        }
    }

    * If no varlist given, read all variables from the index.
    * If vars were given but none exist in the dataset, error -- don't silently exit.
    if "`varlist'" == "" & "`rawvars'" != "" {
        di as err "None of the specified variables exist in the dataset."
        exit 111
    }
    if "`varlist'" == "" & "`rawvars'" == "" {
        tempfile idxtmp
        preserve
        quietly {
            import excel using `"`fname'"', sheet("index") firstrow clear allstring
            capture confirm variable varname
            if _rc {
                restore
                di as err "Sheet [index] missing or malformed"
                exit 111
            }
            save `"`idxtmp'"', replace
        }
        restore

        preserve
        quietly use `"`idxtmp'"', clear
        local nrows = _N
        forval i = 1/`nrows' {
            local varlist `varlist' `=varname[`i']'
        }
        restore
    }

    local any_conflict 0
    local any_missing  0

    di as txt _newline "{hline 60}"
    di as res "  naam check: `fname'"
    di as txt "{hline 60}"

    foreach v of local varlist {

        capture confirm variable `v'
        if _rc {
            di as txt _newline "  `v': not found in dataset -- skipped"
            continue
        }

        * Check variable has a value label attached
        local lbname : value label `v'
        if "`lbname'" == "" {
            di as txt _newline "  `v': no value label attached in dataset"
        }

        * Try to load mapping sheet from Excel
        tempfile valtmp
        local shname = substr("`v'",1,31)
        preserve
        capture quietly {
            import excel using `"`fname'"', sheet("index") firstrow clear allstring
            confirm variable sheetname
            count
            local nidx = r(N)
            forval si = 1/`nidx' {
                if varname[`si'] == "`v'" {
                    local shname = sheetname[`si']
                }
            }
        }
        restore
        local sheet_ok 1
        preserve
        quietly {
            capture {
                import excel using `"`fname'"', ///
                    sheet("`shname'") firstrow clear allstring
                confirm variable numeric_code
                confirm variable string_value
                save `"`valtmp'"', replace
            }
            if _rc local sheet_ok 0
        }
        restore

        if !`sheet_ok' {
            di as txt _newline "  `v': no mapping sheet in Excel file -- skipped"
            continue
        }

        * Load Excel mapping into locals
        preserve
        quietly use `"`valtmp'"', clear
        capture confirm numeric variable numeric_code
        if _rc {
            destring numeric_code, replace
        }
        capture {
            assert !missing(numeric_code)
            isid numeric_code
        }
        if _rc {
            restore
            di as txt _newline "  `v': mapping sheet is invalid or internally inconsistent -- skipped"
            continue
        }
        local nxl = _N
        forval j = 1/`nxl' {
            local xc_`j' = numeric_code[`j']
            local xl_`j' = string_value[`j']
        }
        restore

        * Compare each Excel entry against the in-memory value label
        local n_ok       0
        local n_conflict 0
        local n_missing  0

        di as txt _newline "  Variable: " as res "`v'"
        di as txt "  {hline 50}"

        forval j = 1/`nxl' {
            local code  = `xc_`j''
            local xl_lbl `"`xl_`j''"'

            if "`lbname'" != "" {
                local mem_lbl : label `lbname' `code', strict
            }
            else {
                local mem_lbl ""
            }

            if `"`mem_lbl'"' == "" & `"`xl_lbl'"' != "" {
                di as txt "    code `code': " as txt "not in dataset" ///
                    as txt "  (Excel: `xl_lbl')"
                local ++n_missing
                local any_missing 1
            }
            else if `"`mem_lbl'"' != `"`xl_lbl'"' {
                di as txt "    code `code': " as err "CONFLICT" as txt ///
                    "  dataset=`mem_lbl'  |  Excel=`xl_lbl'"
                local ++n_conflict
                local any_conflict 1
            }
            else {
                local ++n_ok
            }
        }

        * Check for codes in memory not in Excel
        if "`lbname'" != "" {
            quietly label list `lbname'
            local kmin = r(min)
            local kmax = r(max)
            forval code = `kmin'/`kmax' {
                local mem_lbl : label `lbname' `code', strict
                if `"`mem_lbl'"' == "" continue
                local found 0
                forval j = 1/`nxl' {
                    if `xc_`j'' == `code' {
                        local found 1
                    }
                }
                if !`found' {
                    di as txt "    code `code': " as err "in dataset but NOT in Excel" ///
                        as txt "  (dataset: `mem_lbl')"
                    local ++n_conflict
                    local any_conflict 1
                }
            }
        }

        di as txt "  {hline 50}"
        di as txt "    OK: `n_ok'  |  Conflicts: `n_conflict'  |  Missing in dataset: `n_missing'"
    }

    di as txt _newline "{hline 60}"
    if `any_conflict' {
        di as err "  CHECK FAILED: label conflicts detected. Review output above."
    }
    else if `any_missing' {
        di as txt "  CHECK NOTE: some mapping categories are not present in this dataset."
        di as res "  No conflicts -- all codes that exist in the dataset match the saved mapping."
    }
    else {
        di as res "  CHECK PASSED: all labels match the saved mapping."
    }
    di as txt "{hline 60}"
    exit 0
end


* -----------------------------------------------------------------------------
program define naam_compare
* Compare two naam Excel mapping files against each other.
* Reports variables present in one file but not the other, and any
* code-to-label conflicts for variables that appear in both.
* Does not require any dataset in memory.
* -----------------------------------------------------------------------------
    version 14.0
    syntax using/ [, USing2(string)]

    if `"`using2'"' == "" {
        di as err "naam compare requires two filenames: using filename1, using2(filename2)"
        exit 198
    }

    local fname1 `"`using'"'
    local fname2 `"`using2'"'

    if substr(`"`fname1'"',-5,5)!=".xlsx" & substr(`"`fname1'"',-4,4)!=".xls" {
        local fname1 `"`fname1'.xlsx"'
    }
    if substr(`"`fname2'"',-5,5)!=".xlsx" & substr(`"`fname2'"',-4,4)!=".xls" {
        local fname2 `"`fname2'.xlsx"'
    }

    confirm file `"`fname1'"'
    confirm file `"`fname2'"'

    * Read index from file 1
    tempfile idx1 idx2
    preserve
    quietly {
        import excel using `"`fname1'"', sheet("index") firstrow clear allstring
        capture confirm variable varname
        if _rc {
            restore
            di as err "Sheet [index] missing or malformed in `fname1'"
            exit 111
        }
        save `"`idx1'"', replace
    }
    restore

    preserve
    quietly use `"`idx1'"', clear
    capture confirm variable sheetname
    local has_sheetname1 = !_rc
    local n1 = _N
    forval i = 1/`n1' {
        local f1_var_`i' = varname[`i']
        local f1_type_`i' = type[`i']
        if `has_sheetname1' {
            local f1_sheet_`i' = sheetname[`i']
        }
        else {
            local f1_sheet_`i' = substr("`f1_var_`i''",1,31)
        }
    }
    restore

    * Read index from file 2
    preserve
    quietly {
        import excel using `"`fname2'"', sheet("index") firstrow clear allstring
        capture confirm variable varname
        if _rc {
            restore
            di as err "Sheet [index] missing or malformed in `fname2'"
            exit 111
        }
        save `"`idx2'"', replace
    }
    restore

    preserve
    quietly use `"`idx2'"', clear
    capture confirm variable sheetname
    local has_sheetname2 = !_rc
    local n2 = _N
    forval i = 1/`n2' {
        local f2_var_`i' = varname[`i']
        local f2_type_`i' = type[`i']
        if `has_sheetname2' {
            local f2_sheet_`i' = sheetname[`i']
        }
        else {
            local f2_sheet_`i' = substr("`f2_var_`i''",1,31)
        }
    }
    restore

    * Header
    di as txt _newline "{hline 60}"
    di as res "  naam compare"
    di as txt "  File 1: `fname1'"
    di as txt "  File 2: `fname2'"
    di as txt "{hline 60}"

    * Find variables only in file 1
    forval i = 1/`n1' {
        local v "`f1_var_`i''"
        local found 0
        forval j = 1/`n2' {
            if "`f2_var_`j''" == "`v'" local found 1
        }
        if !`found' {
            di as txt _newline "  `v': " as err "only in File 1"
        }
    }

    * Find variables only in file 2
    forval j = 1/`n2' {
        local v "`f2_var_`j''"
        local found 0
        forval i = 1/`n1' {
            if "`f1_var_`i''" == "`v'" local found 1
        }
        if !`found' {
            di as txt _newline "  `v': " as err "only in File 2"
        }
    }

    * Compare variables present in both files
    local any_conflict 0
    forval i = 1/`n1' {
        local v "`f1_var_`i''"
        local shname1 "`f1_sheet_`i''"

        * Check if v is in file 2
        local in2 0
        local shname2 = substr("`v'",1,31)
        forval j = 1/`n2' {
            if "`f2_var_`j''" == "`v'" {
                local in2 1
                local shname2 "`f2_sheet_`j''"
            }
        }
        if !`in2' continue

        * Load mapping from file 1
        tempfile m1 m2
        local s1_ok 1
        preserve
        quietly {
            capture {
                import excel using `"`fname1'"', ///
                    sheet("`shname1'") firstrow clear allstring
                confirm variable numeric_code
                confirm variable string_value
                save `"`m1'"', replace
            }
            if _rc local s1_ok 0
        }
        restore
        if !`s1_ok' continue

        * Load mapping from file 2
        local s2_ok 1
        preserve
        quietly {
            capture {
                import excel using `"`fname2'"', ///
                    sheet("`shname2'") firstrow clear allstring
                confirm variable numeric_code
                confirm variable string_value
                save `"`m2'"', replace
            }
            if _rc local s2_ok 0
        }
        restore
        if !`s2_ok' continue

        * Load into locals
        preserve
        quietly use `"`m1'"', clear
        local nv1 = _N
        forval j = 1/`nv1' {
            local m1c_`j' = numeric_code[`j']
            local m1l_`j' = string_value[`j']
        }
        restore

        preserve
        quietly use `"`m2'"', clear
        local nv2 = _N
        forval j = 1/`nv2' {
            local m2c_`j' = numeric_code[`j']
            local m2l_`j' = string_value[`j']
        }
        restore

        * Compare file 1 entries against file 2
        local n_ok      0
        local n_conf    0
        local hdr_shown 0

        forval j = 1/`nv1' {
            local code `m1c_`j''
            local lbl1 `"`m1l_`j''"'
            local lbl2 ""
            forval k = 1/`nv2' {
                if `m2c_`k'' == `code' {
                    local lbl2 `"`m2l_`k''"'
                }
            }
            if `"`lbl2'"' == "" {
                if !`hdr_shown' {
                    di as txt _newline "  Variable: " as res "`v'"
                    di as txt "  {hline 50}"
                    local hdr_shown 1
                }
                di as txt "    code `code': " as err "only in File 1" ///
                    as txt "  (`lbl1')"
                local ++n_conf
                local any_conflict 1
            }
            else if `"`lbl1'"' != `"`lbl2'"' {
                if !`hdr_shown' {
                    di as txt _newline "  Variable: " as res "`v'"
                    di as txt "  {hline 50}"
                    local hdr_shown 1
                }
                di as txt "    code `code': " as err "CONFLICT" as txt ///
                    "  File1=`lbl1'  |  File2=`lbl2'"
                local ++n_conf
                local any_conflict 1
            }
            else {
                local ++n_ok
            }
        }

        * Codes in file 2 not in file 1
        forval k = 1/`nv2' {
            local code `m2c_`k''
            local lbl2 `"`m2l_`k''"'
            local found 0
            forval j = 1/`nv1' {
                if `m1c_`j'' == `code' local found 1
            }
            if !`found' {
                if !`hdr_shown' {
                    di as txt _newline "  Variable: " as res "`v'"
                    di as txt "  {hline 50}"
                    local hdr_shown 1
                }
                di as txt "    code `code': " as err "only in File 2" ///
                    as txt "  (`lbl2')"
                local ++n_conf
                local any_conflict 1
            }
        }

        if `hdr_shown' {
            di as txt "  {hline 50}"
            di as txt "    OK: `n_ok'  |  Conflicts / missing: `n_conf'"
        }
    }

    di as txt _newline "{hline 60}"
    if `any_conflict' {
        di as err "  COMPARE FAILED: differences found. Review output above."
    }
    else {
        di as res "  COMPARE PASSED: both files are fully consistent."
    }
    di as txt "{hline 60}"
end
