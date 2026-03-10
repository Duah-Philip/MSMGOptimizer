## Resubmission
Addressed all issues raised in the CRAN review:
* Expanded all acronyms in DESCRIPTION and documentation (MSMG, LCI, LCA, CSV, NSF)
* Added single quotes around all software names ('Shiny', 'SimaPro', 'Excel', 'Python', 'RStudio')
* Added \value tag to all exported functions
* Replaced \dontrun{} with if(interactive()){}
* Added executable code chunks to vignette
* Added reference to Pizzol (2022) in DESCRIPTION
* Fixed non-portable file names (spaces replaced with underscores)

## Win-builder results (R-devel)
0 errors | 0 warnings | 1 note

## Note
* New submission
* Possibly misspelled words: LCA, LCI, MSMG, Pizzol — these are
  intentional technical terms and proper nouns, all expanded in
  the Description field
