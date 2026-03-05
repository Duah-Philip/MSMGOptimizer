# app.R - MSMG SimaPro CSV Optimizer



library(shiny)
library(shinydashboard)
library(readxl)
library(dplyr)
library(DT)
library(waiter)
library(htmltools)
library(zip)



# ─────────────────────────────────────────────────────────────────────────────
# CONVERSION FUNCTION
# ─────────────────────────────────────────────────────────────────────────────
to_spcsv <- function(dataframe, file_path) {
  debug_info <- list()
  debug_log <- function(msg) {
    debug_info[[length(debug_info) + 1]] <<- msg
    print(paste("DEBUG:", msg))
  }
  debug_log(paste("Dataframe dimensions:", nrow(dataframe), "rows,", ncol(dataframe), "columns"))
  con <- NULL
  on.exit({
    if (!is.null(con) && isOpen(con)) {
      close(con)
      debug_log("Connection closed properly")
    }
    if (exists("error_occurred") && error_occurred) {
      writeLines(paste(unlist(debug_info), collapse = "\n"), paste0(file_path, ".debug.log"))
    }
  })
  error_occurred <- FALSE
  if (nrow(dataframe) < 3) { debug_log("ERROR: Input data has too few rows"); stop("Input data must have at least 3 rows") }
  if (ncol(dataframe) < 3) { debug_log("ERROR: Input data has too few columns"); stop("Input data must have at least 3 columns") }
  debug_log("First row content:")
  if (nrow(dataframe) > 0) debug_log(paste(as.character(dataframe[1,]), collapse = " | "))
  tryCatch({
    con <- file(file_path, "w")
    writeLines("{CSV separator: Semicolon}", con)
    writeLines("{CSV Format version: 7.0.0}", con)
    writeLines("{Decimal separator: .}", con)
    writeLines("{Date separator: /}", con)
    writeLines("{Short date format: dd/MM/yyyy}", con)
    writeLines("", con)
    debug_log(paste("File opened successfully:", file_path))
  }, error = function(e) {
    debug_log(paste("Error opening file:", e$message))
    error_occurred <- TRUE
    stop(paste("Failed to create output file:", e$message))
  })
  fields <- c("Process","Category type","Time Period","Geography","Technology","Representativeness",
              "Multiple output allocation","Substitution allocation","Cut off rules","Capital goods",
              "Boundary with nature","Record","Generator","Literature references","Collection method",
              "Data treatment","Verification","Products","Materials/fuels","Resources",
              "Emissions to air","Emissions to water","Emissions to soil","Final waste flows",
              "Non material emission","Social issues","Economic issues","Waste to treatment","End")
  fields_value <- list("","","Unspecified","Unspecified","Unspecified","Unspecified","Unspecified",
                       "Unspecified","Unspecified","Unspecified","Unspecified","","","","","",
                       "Comment","",list(),list(),list(),list(),list(),list(),"",list(),list(),list(),"")
  dataframe[is.na(dataframe)] <- ""
  default_process <- FALSE
  if (ncol(dataframe) < 5) {
    debug_log("WARNING: Excel has fewer than 5 columns, using default process")
    default_process <- TRUE
    dataframe <- cbind(dataframe, Default_Process = rep("", nrow(dataframe)))
    if (nrow(dataframe) > 0) dataframe[1, ncol(dataframe)] <- "Default_Process"
    debug_log(paste("New dataframe dimensions after adding default process:", nrow(dataframe), "rows,", ncol(dataframe), "columns"))
  }
  tryCatch({
    if (nrow(dataframe) >= 1) {
      if (default_process) {
        processes <- dataframe[1, ncol(dataframe), drop=FALSE]
        debug_log("Using default process name")
      } else if (ncol(dataframe) >= 5) {
        processes <- dataframe[1, 5:ncol(dataframe), drop=FALSE]
        debug_log(paste("Found", ncol(processes), "potential processes"))
      } else {
        processes <- data.frame(Default="Default_Process")
        debug_log("Excel format incorrect: Using default process name")
      }
      if (!default_process) {
        valid_indices <- which(!is.na(processes) & processes != "", arr.ind = TRUE)
        debug_log(paste("Valid process indices:", paste(valid_indices, collapse=", ")))
        if (length(valid_indices) == 0) {
          debug_log("WARNING: No valid processes found, using default")
          processes <- data.frame(Default="Default_Process")
          default_process <- TRUE
        }
      }
    } else {
      debug_log("ERROR: Empty dataframe")
      stop("Dataframe has no rows")
    }
  }, error = function(e) {
    debug_log(paste("Error identifying processes:", e$message))
    error_occurred <- TRUE
    stop(paste("Failed to identify processes:", e$message))
  })
  debug_log(paste("Starting process processing, with", length(processes), "potential processes"))
  process_count <- if(is.data.frame(processes)) ncol(processes) else length(processes)
  debug_log(paste("Process count:", process_count))
  for (i in 1:process_count) {
    tryCatch({
      col_idx <- if(default_process) ncol(dataframe) else (i + 4)
      debug_log(paste("Processing column index:", col_idx))
      if (col_idx > ncol(dataframe)) { debug_log(paste("Column index", col_idx, "is out of bounds, skipping")); next }
      if (ncol(dataframe) >= col_idx) {
        process <- dataframe[, col_idx, drop=FALSE]
        debug_log(paste("Extracted process column with", nrow(process), "rows"))
      } else {
        debug_log(paste("Cannot extract column", col_idx, "- out of bounds")); next
      }
      fields_value[[2]] <- ""
      if (nrow(dataframe) >= 6 && ncol(dataframe) >= col_idx) {
        category_type <- as.character(dataframe[6, col_idx])
        debug_log(paste("Category type found:", category_type))
        if (!is.na(category_type)) fields_value[[2]] <- category_type
      } else {
        debug_log("Not enough rows for category type or column doesn't exist")
      }
      product_name <- "Default_Product"; product_unit <- "kg"; product_amount <- "1"; product_comment <- ""
      if (nrow(dataframe) >= 1 && ncol(dataframe) >= col_idx) { temp <- as.character(dataframe[1, col_idx]); if (!is.na(temp) && temp != "") product_name <- temp }
      if (nrow(dataframe) >= 2 && ncol(dataframe) >= col_idx) { temp <- as.character(dataframe[2, col_idx]); if (!is.na(temp) && temp != "") product_unit <- temp }
      if (nrow(dataframe) >= 3 && ncol(dataframe) >= col_idx) { temp <- as.character(dataframe[3, col_idx]); if (!is.na(temp) && temp != "") product_amount <- temp }
      if (nrow(dataframe) >= 5 && ncol(dataframe) >= col_idx) { temp <- as.character(dataframe[5, col_idx]); if (!is.na(temp)) product_comment <- temp }
      debug_log(paste("Product info:", product_name, product_unit, product_amount, product_comment))
      products <- sprintf('"%s";"%s";"%s";"100%%";"not defined";"%s"', product_name, product_unit, product_amount, product_comment)
      fields_value[[18]] <- products
      matfuel_list <- c(); raw_list <- c(); air_list <- c(); water_list <- c(); soil_list <- c()
      finalwaste_list <- c(); social_list <- c(); economic_list <- c(); wastetotreatment_list <- c()
      debug_log("Initialized exchange type lists")
      start_row <- min(7, nrow(dataframe))
      if (start_row < 7) debug_log(paste("WARNING: Using start row", start_row, "instead of 7"))
      if (start_row <= nrow(dataframe)) {
        debug_log(paste("Processing exchanges from row", start_row, "to", nrow(dataframe)))
        for (j in start_row:nrow(dataframe)) {
          tryCatch({
            exchange_type <- ""; if (ncol(dataframe) >= 1) { temp <- dataframe[j, 1]; if (!is.na(temp)) exchange_type <- as.character(temp) }
            exchange_value <- ""; if (ncol(dataframe) >= col_idx) { temp <- dataframe[j, col_idx]; if (!is.na(temp)) exchange_value <- as.character(temp) }
            debug_log(paste("Row", j, "- Exchange type:", exchange_type, "Value:", exchange_value))
            if (exchange_value != "") {
              exchange_name <- "Unknown"; if (ncol(dataframe) >= 2) { temp <- dataframe[j, 2]; if (!is.na(temp)) exchange_name <- as.character(temp) }
              exchange_unit <- "kg"; if (ncol(dataframe) >= 4) { temp <- dataframe[j, 4]; if (!is.na(temp)) exchange_unit <- as.character(temp) }
              debug_log(paste("Exchange details - Name:", exchange_name, "Unit:", exchange_unit))
              if (exchange_type == "") {
                matfuel_list <- c(matfuel_list, sprintf('"%s";"%s";"%s";"Undefined";0;0;0', exchange_name, exchange_unit, exchange_value))
              } else if (tolower(exchange_type) == "raw") {
                raw_list <- c(raw_list, sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0', exchange_name, "", exchange_unit, exchange_value))
              } else if (tolower(exchange_type) == "air") {
                air_list <- c(air_list, sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0', exchange_name, "", exchange_unit, exchange_value))
              } else if (tolower(exchange_type) == "water") {
                water_list <- c(water_list, sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0', exchange_name, "", exchange_unit, exchange_value))
              } else if (tolower(exchange_type) == "soil") {
                soil_list <- c(soil_list, sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0', exchange_name, "", exchange_unit, exchange_value))
              } else if (tolower(exchange_type) == "waste") {
                finalwaste_list <- c(finalwaste_list, sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0', exchange_name, "", exchange_unit, exchange_value))
              } else if (tolower(exchange_type) == "social") {
                social_list <- c(social_list, sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0', exchange_name, "", exchange_unit, exchange_value))
              } else if (tolower(exchange_type) == "economic") {
                economic_list <- c(economic_list, sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0', exchange_name, "", exchange_unit, exchange_value))
              } else if (tolower(exchange_type) == "wastetotreatment") {
                wastetotreatment_list <- c(wastetotreatment_list, sprintf('"%s";"%s";"%s";"Undefined";0;0;0', exchange_name, exchange_unit, exchange_value))
              } else {
                debug_log(paste("Unknown exchange type:", exchange_type, "- skipping"))
              }
            }
          }, error = function(e) { debug_log(paste("Error processing row", j, ":", e$message)) })
        }
      } else { debug_log("No rows to process for exchanges") }
      fields_value[[19]] <- matfuel_list; fields_value[[20]] <- raw_list; fields_value[[21]] <- air_list
      fields_value[[22]] <- water_list; fields_value[[23]] <- soil_list; fields_value[[24]] <- finalwaste_list
      fields_value[[26]] <- social_list; fields_value[[27]] <- economic_list; fields_value[[28]] <- wastetotreatment_list
      debug_log(paste("Exchange counts - Materials/fuels:", length(matfuel_list), "Resources:", length(raw_list),
                      "Air:", length(air_list), "Water:", length(water_list), "Soil:", length(soil_list),
                      "Waste:", length(finalwaste_list), "Social:", length(social_list),
                      "Economic:", length(economic_list), "Waste treatment:", length(wastetotreatment_list)))
      tryCatch({
        debug_log("Writing fields to output file")
        for (el in 1:length(fields)) {
          writeLines(fields[el], con)
          if (is.atomic(fields_value[[el]]) && length(fields_value[[el]]) <= 1) {
            writeLines(as.character(fields_value[[el]]), con)
          } else {
            if (length(fields_value[[el]]) > 0) {
              for (item in fields_value[[el]]) writeLines(as.character(item), con)
            } else { writeLines("", con) }
          }
          writeLines("", con)
        }
        debug_log("Successfully completed writing process")
      }, error = function(e) { debug_log(paste("Error writing to output file:", e$message)); error_occurred <- TRUE })
    }, error = function(e) { debug_log(paste("Process loop error:", e$message)); error_occurred <- TRUE })
  }
  if (exists("error_occurred") && error_occurred) {
    writeLines(paste(unlist(debug_info), collapse = "\n"), paste0(file_path, ".debug.log"))
    warning(paste("Errors occurred during conversion. Debug log written to", paste0(file_path, ".debug.log")))
  }
  return(!exists("error_occurred") || !error_occurred)
}

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS — Theme
# ─────────────────────────────────────────────────────────────────────────────
custom_css <- "
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;1,9..40,300&display=swap');

/* ── Root variables — MSMG Brand Palette ── */
:root {
  --navy:         #142348;
  --navy-mid:     #1e3266;
  --navy-soft:    #2a4080;
  --ink-soft:     #3a4a6a;
  --ink-muted:    #7a8fa6;
  --surface:      #f2f4f8;
  --surface-2:    #e8ecf4;
  --white:        #ffffff;
  --lime:         #9dc426;
  --lime-bright:  #b8e032;
  --lime-light:   #eef7c8;
  --lime-pale:    #f6fbea;
  --gold:         #c8922a;
  --gold-light:   #fdf3e0;
  --border:       #d4dae8;
  --danger:       #c0392b;
  --radius:       10px;
  --shadow-sm:    0 2px 8px rgba(20,35,72,0.08);
  --shadow-md:    0 6px 24px rgba(20,35,72,0.12);
  --shadow-lg:    0 16px 48px rgba(20,35,72,0.16);
  --transition:   all 0.22s cubic-bezier(0.4,0,0.2,1);

  /* Aliases for component use */
  --accent:       var(--lime);
  --accent-2:     var(--lime-bright);
  --accent-light: var(--lime-light);
  --ink:          var(--navy);
}

/* ── Global reset ── */
* { box-sizing: border-box; }
body {
  font-family: 'DM Sans', sans-serif;
  background: var(--surface) !important;
  color: var(--ink) !important;
  font-size: 14.5px;
  line-height: 1.65;
}

/* ── Sidebar ── */
.main-sidebar, .left-side {
  background: var(--navy) !important;
  border-right: none !important;
  box-shadow: 4px 0 24px rgba(20,35,72,0.28) !important;
}
.sidebar-menu > li > a {
  font-family: 'DM Sans', sans-serif !important;
  font-weight: 500 !important;
  font-size: 13.5px !important;
  letter-spacing: 0.3px !important;
  color: rgba(255,255,255,0.55) !important;
  padding: 13px 20px 13px 22px !important;
  border-left: 3px solid transparent !important;
  transition: var(--transition) !important;
}
.sidebar-menu > li > a:hover,
.sidebar-menu > li.active > a {
  color: var(--navy) !important;
  background: var(--lime) !important;
  border-left: 3px solid var(--lime-bright) !important;
}
.sidebar-menu > li > a > .fa {
  width: 22px;
  color: var(--lime) !important;
  font-size: 14px !important;
  margin-right: 10px !important;
}
.sidebar-menu > li.active > a > .fa,
.sidebar-menu > li > a:hover > .fa { color: var(--navy) !important; }

/* Sidebar logo block */
.logo-block {
  padding: 22px 20px 18px 20px;
  border-bottom: 1px solid rgba(157,196,38,0.2);
  margin-bottom: 8px;
}
.logo-block img {
  width: 110px;
  height: 110px;
  display: block;
  margin: 0 auto 12px auto;
  filter: drop-shadow(0 3px 12px rgba(0,0,0,0.5));
}
.logo-block .app-title {
  font-family: 'DM Serif Display', serif !important;
  color: #ffffff !important;
  font-size: 14px !important;
  text-align: center;
  line-height: 1.4;
  letter-spacing: 0.2px;
}
.logo-block .app-subtitle {
  color: var(--lime) !important;
  font-size: 10.5px !important;
  text-align: center;
  letter-spacing: 1.2px;
  text-transform: uppercase;
  margin-top: 4px;
  font-weight: 600;
}

/* ── Header ── */
.main-header .navbar,
.main-header .logo {
  background: var(--white) !important;
  border-bottom: 1px solid var(--border) !important;
  box-shadow: var(--shadow-sm) !important;
}
.main-header .logo {
  background: var(--navy) !important;
  font-family: 'DM Serif Display', serif !important;
  font-size: 16px !important;
  color: #ffffff !important;
  letter-spacing: 0.3px !important;
  border-bottom: none !important;
}
.main-header .navbar-custom-menu > .navbar-nav > li > a {
  color: var(--ink-soft) !important;
}
.main-header .sidebar-toggle {
  color: var(--ink-soft) !important;
  font-size: 16px !important;
}

/* ── Content wrapper ── */
.content-wrapper, .right-side {
  background: var(--surface) !important;
  padding-top: 0 !important;
}
.content { padding: 24px 28px !important; }

/* ── Page hero banner ── */
.page-hero {
  background: linear-gradient(135deg, var(--navy) 0%, var(--navy-mid) 55%, #0e2a14 100%);
  border-radius: var(--radius);
  padding: 32px 36px;
  margin-bottom: 24px;
  position: relative;
  overflow: hidden;
}
.page-hero::before {
  content: '';
  position: absolute;
  top: -40px; right: -40px;
  width: 220px; height: 220px;
  border-radius: 50%;
  background: rgba(157,196,38,0.12);
  pointer-events: none;
}
.page-hero::after {
  content: '';
  position: absolute;
  bottom: -60px; left: 30%;
  width: 300px; height: 300px;
  border-radius: 50%;
  background: rgba(157,196,38,0.06);
  pointer-events: none;
}
.page-hero h2 {
  font-family: 'DM Serif Display', serif;
  color: #ffffff;
  font-size: 24px;
  margin: 0 0 6px 0;
  font-weight: 400;
}
.page-hero p {
  color: rgba(255,255,255,0.65);
  font-size: 13.5px;
  margin: 0;
  font-weight: 300;
  max-width: 600px;
}
.hero-badge {
  display: inline-block;
  background: rgba(157,196,38,0.22);
  color: var(--lime-bright);
  border: 1px solid rgba(157,196,38,0.55);
  border-radius: 20px;
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.8px;
  text-transform: uppercase;
  padding: 3px 12px;
  margin-bottom: 12px;
}

/* ── Stat cards ── */
.stat-cards { display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; }
.stat-card {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 18px 22px;
  flex: 1; min-width: 160px;
  box-shadow: var(--shadow-sm);
  position: relative;
  overflow: hidden;
  transition: var(--transition);
}
.stat-card:hover { transform: translateY(-2px); box-shadow: var(--shadow-md); }
.stat-card::before {
  content: '';
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 3px;
  background: linear-gradient(90deg, var(--lime), var(--lime-bright));
  border-radius: var(--radius) var(--radius) 0 0;
}
.stat-card .stat-icon {
  width: 36px; height: 36px;
  background: var(--lime-light);
  border-radius: 8px;
  display: flex; align-items: center; justify-content: center;
  margin-bottom: 10px;
  font-size: 15px;
  color: var(--navy);
}
.stat-card .stat-value {
  font-family: 'DM Serif Display', serif;
  font-size: 26px;
  color: var(--navy);
  line-height: 1;
  margin-bottom: 4px;
}
.stat-card .stat-label {
  font-size: 12px;
  color: var(--ink-muted);
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}

/* ── Custom cards ── */
.sci-card {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow-sm);
  margin-bottom: 22px;
  overflow: hidden;
  transition: var(--transition);
}
.sci-card:hover { box-shadow: var(--shadow-md); }
.sci-card-header {
  padding: 16px 22px;
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: center;
  gap: 10px;
  background: var(--white);
}
.sci-card-header .card-icon {
  width: 30px; height: 30px;
  border-radius: 7px;
  display: flex; align-items: center; justify-content: center;
  font-size: 13px;
  flex-shrink: 0;
}
.icon-green { background: var(--lime-light); color: var(--navy); }
.icon-blue  { background: #daeaf8; color: #1a6b9e; }
.icon-amber { background: var(--gold-light); color: var(--gold); }
.sci-card-header h4 {
  font-family: 'DM Sans', sans-serif;
  font-weight: 600;
  font-size: 14.5px;
  color: var(--ink);
  margin: 0;
  letter-spacing: 0.1px;
}
.sci-card-body { padding: 22px; }

/* ── File input ── */
.form-group label {
  font-weight: 600 !important;
  font-size: 13px !important;
  color: var(--ink-soft) !important;
  letter-spacing: 0.2px !important;
  text-transform: uppercase !important;
  margin-bottom: 8px !important;
}
.form-control {
  border: 1.5px solid var(--border) !important;
  border-radius: 8px !important;
  font-family: 'DM Sans', sans-serif !important;
  font-size: 14px !important;
  color: var(--ink) !important;
  background: var(--white) !important;
  transition: var(--transition) !important;
  padding: 9px 14px !important;
  height: auto !important;
}
.form-control:focus {
  border-color: var(--lime) !important;
  box-shadow: 0 0 0 3px rgba(157,196,38,0.2) !important;
  outline: none !important;
}

/* Custom file input */
.btn-file {
  background: var(--white) !important;
  border: 1.5px dashed var(--border) !important;
  color: var(--ink-soft) !important;
  border-radius: 8px !important;
  font-family: 'DM Sans', sans-serif !important;
  font-size: 13.5px !important;
  transition: var(--transition) !important;
  padding: 10px 18px !important;
}
.btn-file:hover {
  border-color: var(--lime) !important;
  color: var(--navy) !important;
  background: var(--lime-pale) !important;
}
input[type='file'] { display: none !important; }

/* ── Action buttons ── */
.btn-convert {
  background: linear-gradient(135deg, var(--lime) 0%, var(--lime-bright) 100%) !important;
  color: var(--navy) !important;
  border: none !important;
  border-radius: 8px !important;
  font-family: 'DM Sans', sans-serif !important;
  font-weight: 600 !important;
  font-size: 14px !important;
  letter-spacing: 0.3px !important;
  padding: 11px 28px !important;
  cursor: pointer !important;
  box-shadow: 0 4px 14px rgba(157,196,38,0.45) !important;
  transition: var(--transition) !important;
  display: inline-flex !important;
  align-items: center !important;
  gap: 8px !important;
}
.btn-convert:hover {
  transform: translateY(-1px) !important;
  box-shadow: 0 6px 20px rgba(157,196,38,0.55) !important;
  background: linear-gradient(135deg, #7da01a 0%, var(--lime) 100%) !important;
}
.btn-convert:active { transform: translateY(0) !important; }

.btn-download {
  background: var(--white) !important;
  color: var(--navy) !important;
  border: 1.5px solid var(--lime) !important;
  border-radius: 8px !important;
  font-family: 'DM Sans', sans-serif !important;
  font-weight: 600 !important;
  font-size: 14px !important;
  letter-spacing: 0.3px !important;
  padding: 10px 24px !important;
  cursor: pointer !important;
  transition: var(--transition) !important;
  display: inline-flex !important;
  align-items: center !important;
  gap: 8px !important;
}
.btn-download:hover {
  background: var(--lime-light) !important;
  border-color: var(--lime-bright) !important;
  transform: translateY(-1px) !important;
}

/* ── Button row ── */
.action-row {
  display: flex;
  gap: 12px;
  flex-wrap: wrap;
  align-items: center;
  padding-top: 8px;
}

/* ── Status log ── */
.status-log {
  background: #0f1923 !important;
  color: #7aefb2 !important;
  font-family: 'DM Mono', 'Fira Code', 'Courier New', monospace !important;
  font-size: 12.5px !important;
  line-height: 1.7 !important;
  border-radius: 8px !important;
  border: none !important;
  padding: 16px 18px !important;
  min-height: 120px !important;
  box-shadow: inset 0 2px 8px rgba(0,0,0,0.3) !important;
  white-space: pre-wrap !important;
  resize: vertical !important;
}
pre.shiny-text-output {
  background: #0f1923 !important;
  color: #7aefb2 !important;
  font-family: 'Courier New', monospace !important;
  font-size: 12.5px !important;
  border: none !important;
  border-radius: 8px !important;
  padding: 16px 18px !important;
  min-height: 120px !important;
  box-shadow: inset 0 2px 8px rgba(0,0,0,0.3) !important;
}

/* ── DataTable ── */
.dataTables_wrapper { font-family: 'DM Sans', sans-serif !important; }
table.dataTable thead th {
  background: var(--navy) !important;
  color: #ffffff !important;
  font-size: 12.5px !important;
  font-weight: 600 !important;
  letter-spacing: 0.4px !important;
  text-transform: uppercase !important;
  border: none !important;
  padding: 12px 14px !important;
}
table.dataTable tbody tr {
  background: var(--white) !important;
  transition: background 0.15s ease !important;
}
table.dataTable tbody tr:hover { background: var(--lime-pale) !important; }
table.dataTable tbody tr:nth-child(even) { background: #f6f8fc !important; }
table.dataTable tbody tr:nth-child(even):hover { background: var(--lime-pale) !important; }
table.dataTable tbody td {
  border-color: var(--border) !important;
  font-size: 13.5px !important;
  padding: 10px 14px !important;
  color: var(--ink-soft) !important;
}
.dataTables_length select,
.dataTables_filter input {
  border: 1.5px solid var(--border) !important;
  border-radius: 6px !important;
  font-family: 'DM Sans', sans-serif !important;
  font-size: 13px !important;
  padding: 5px 10px !important;
}
.dataTables_paginate .paginate_button {
  font-family: 'DM Sans', sans-serif !important;
  font-size: 13px !important;
  border-radius: 6px !important;
}
.dataTables_paginate .paginate_button.current,
.dataTables_paginate .paginate_button.current:hover {
  background: var(--navy) !important;
  color: #ffffff !important;
  border-color: var(--navy) !important;
}

/* ── Sheet selector ── */
.selectize-control .selectize-input {
  border: 1.5px solid var(--border) !important;
  border-radius: 8px !important;
  font-family: 'DM Sans', sans-serif !important;
  font-size: 14px !important;
  padding: 9px 14px !important;
  box-shadow: none !important;
  transition: var(--transition) !important;
}
.selectize-control .selectize-input.focus {
  border-color: var(--lime) !important;
  box-shadow: 0 0 0 3px rgba(157,196,38,0.2) !important;
}
.selectize-dropdown { border-radius: 8px !important; border: 1.5px solid var(--border) !important; box-shadow: var(--shadow-md) !important; font-family: 'DM Sans', sans-serif !important; }
.selectize-dropdown .option:hover, .selectize-dropdown .option.selected { background: var(--lime-light) !important; color: var(--navy) !important; }

/* ── About & Instructions ── */
.content-section { margin-bottom: 28px; }
.content-section h3 {
  font-family: 'DM Serif Display', serif;
  font-size: 22px;
  color: var(--navy);
  margin-bottom: 6px;
  font-weight: 400;
}
.content-section h4 {
  font-family: 'DM Sans', sans-serif;
  font-size: 13.5px;
  font-weight: 700;
  color: var(--ink);
  text-transform: uppercase;
  letter-spacing: 0.6px;
  margin: 22px 0 10px 0;
  padding-bottom: 6px;
  border-bottom: 2px solid var(--lime-light);
}
.content-section p, .content-section li {
  color: var(--ink-soft);
  font-size: 14px;
  line-height: 1.75;
}
.content-section ul, .content-section ol {
  padding-left: 20px;
}
.content-section li { margin-bottom: 6px; }
.content-section a {
  color: var(--navy) !important;
  text-decoration: none !important;
  font-weight: 600 !important;
  border-bottom: 1px solid rgba(157,196,38,0.5) !important;
  transition: var(--transition) !important;
}
.content-section a:hover {
  color: var(--navy-soft) !important;
  border-bottom-color: var(--lime) !important;
}

/* Step indicator for instructions */
.step-list { list-style: none !important; padding-left: 0 !important; counter-reset: steps; }
.step-list > li {
  counter-increment: steps;
  position: relative;
  padding-left: 52px;
  margin-bottom: 18px;
  color: var(--ink-soft);
  font-size: 14px;
  line-height: 1.7;
}
.step-list > li::before {
  content: counter(steps);
  position: absolute;
  left: 0; top: 0;
  width: 32px; height: 32px;
  background: linear-gradient(135deg, var(--lime), var(--lime-bright));
  color: var(--navy);
  border-radius: 50%;
  font-family: 'DM Serif Display', serif;
  font-size: 14px;
  display: flex; align-items: center; justify-content: center;
  box-shadow: 0 2px 8px rgba(157,196,38,0.4);
}

/* ── Citation box ── */
.citation-box {
  background: var(--lime-pale);
  border: 1px solid rgba(157,196,38,0.25);
  border-left: 4px solid var(--lime);
  border-radius: 0 8px 8px 0;
  padding: 14px 18px;
  font-size: 13.5px;
  color: var(--ink-soft);
  margin-top: 20px;
  line-height: 1.6;
}
.citation-box strong { color: var(--ink); }

/* ── Info alert ── */
.info-alert {
  background: var(--lime-pale);
  border: 1px solid rgba(157,196,38,0.3);
  border-left: 4px solid var(--lime);
  border-radius: 0 8px 8px 0;
  padding: 12px 16px;
  font-size: 13px;
  color: var(--navy);
  margin-bottom: 16px;
  display: flex;
  align-items: flex-start;
  gap: 10px;
}
.info-alert .fa { margin-top: 2px; color: var(--lime); }

/* ── Footer ── */
.app-footer {
  margin-top: 32px;
  padding: 18px 0 8px 0;
  border-top: 1px solid var(--border);
  text-align: center;
  font-size: 12px;
  color: var(--ink-muted);
  font-weight: 400;
}
.app-footer a { color: var(--lime) !important; text-decoration: none !important; }

/* ── Waiter overlay ── */
.waiter-overlay { border-radius: var(--radius) !important; }
.waiter-overlay-content h4 {
  font-family: 'DM Serif Display', serif !important;
  color: #ffffff !important;
  font-size: 20px !important;
  margin-top: 16px !important;
}
.waiter-overlay-content p {
  color: rgba(255,255,255,0.7) !important;
  font-size: 13px !important;
}

/* ── Responsive ── */
@media (max-width: 768px) {
  .stat-cards { flex-direction: column; }
  .action-row { flex-direction: column; align-items: stretch; }
  .page-hero { padding: 22px 18px; }
  .content { padding: 16px !important; }
}
"

# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────
ui <- dashboardPage(
  skin = "black",

  dashboardHeader(
    title = tags$span(
      style = "font-family: 'DM Serif Display', serif; font-size: 15px; letter-spacing: 0.3px;",
      "MSMG Optimizer"
    ),
    titleWidth = 280
  ),

  dashboardSidebar(
    width = 280,
    tags$div(
      class = "logo-block",

      tags$img(
        src = "Final_MSMG_hex.png",
        alt  = "MSMG Logo",
        onerror = "this.style.display='none'"
      ),
      tags$div(class = "app-title",  "MSMG SimaPro CSV Optimizer"),
      tags$div(class = "app-subtitle", "v 0.1.0  ·  LCA Tools")
    ),
    sidebarMenu(
      useWaiter(),
      menuItem("Converter",    tabName = "home",         icon = icon("sliders-h")),
      menuItem("About",        tabName = "about",        icon = icon("info-circle")),
      menuItem("Instructions", tabName = "instructions", icon = icon("book-open"))
    ),
    tags$div(
      style = "position: absolute; bottom: 16px; left: 0; right: 0; padding: 0 20px;",
      tags$div(
        style = "font-size: 11px; color: rgba(255,255,255,0.28); text-align: center; line-height: 1.6;",
        "NSF Project No. 2219086", tags$br(),
        HTML("&copy; 2025 Philip Duah")
      )
    )
  ),

  dashboardBody(
    tags$head(
      tags$style(HTML(custom_css)),
      tags$link(rel = "stylesheet",
                href = "https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;1,9..40,300&display=swap")
    ),

    tabItems(

      # ── HOME / CONVERTER TAB ──────────────────────────────────────────────
      tabItem(tabName = "home",

              # Hero banner
              tags$div(class = "page-hero",
                       tags$div(class = "hero-badge", tags$i(class="fa fa-leaf", style="margin-right:5px;"), "Life Cycle Assessment"),
                       tags$h2("SimaPro CSV Converter"),
                       tags$p("Transform structured Excel-based LCI data into fully compatible SimaPro CSV format in one click.")
              ),

              # Stat cards (dynamic)
              tags$div(class = "stat-cards",
                       tags$div(class = "stat-card",
                                tags$div(class = "stat-icon", tags$i(class="fa fa-table")),
                                tags$div(class = "stat-value", id = "stat-sheets", "—"),
                                tags$div(class = "stat-label", "Sheets Detected")
                       ),
                       tags$div(class = "stat-card",
                                tags$div(class = "stat-icon", tags$i(class="fa fa-check-circle")),
                                tags$div(class = "stat-value", id = "stat-csvs", "—"),
                                tags$div(class = "stat-label", "CSVs Generated")
                       ),
                       tags$div(class = "stat-card",
                                tags$div(class = "stat-icon", tags$i(class="fa fa-file-archive")),
                                tags$div(class = "stat-value", "ZIP", style="font-size:20px;"),
                                tags$div(class = "stat-label", "Output Format")
                       )
              ),

              # Upload & Convert card
              tags$div(class = "sci-card",
                       tags$div(class = "sci-card-header",
                                tags$div(class = "card-icon icon-green", tags$i(class = "fa fa-upload")),
                                tags$h4("Upload & Convert")
                       ),
                       tags$div(class = "sci-card-body",
                                tags$div(class = "info-alert",
                                         tags$i(class = "fa fa-info-circle"),
                                         tags$span("Accepted formats: ", tags$strong(".xlsx"), " and ", tags$strong(".xls"),
                                                   ". Each worksheet will be converted to an individual SimaPro CSV file.")
                                ),
                                fileInput("file",
                                          label    = NULL,
                                          placeholder = "No file selected — click to browse",
                                          accept   = c(".xlsx", ".xls"),
                                          buttonLabel = tags$span(tags$i(class="fa fa-folder-open", style="margin-right:6px;"), "Browse")),
                                tags$div(class = "action-row",
                                         actionButton("convert", "Convert to SimaPro CSV",
                                                      icon  = icon("exchange-alt"),
                                                      class = "btn-convert"),
                                         downloadButton("downloadZip",
                                                        label = tags$span(tags$i(class="fa fa-download", style="margin-right:6px;"), "Download ZIP"),
                                                        class = "btn-download")
                                )
                       )
              ),

              # Status log card
              tags$div(class = "sci-card",
                       tags$div(class = "sci-card-header",
                                tags$div(class = "card-icon icon-blue", tags$i(class = "fa fa-terminal")),
                                tags$h4("Processing Log")
                       ),
                       tags$div(class = "sci-card-body", style = "padding: 0;",
                                verbatimTextOutput("processingStatus")
                       )
              ),

              # Preview card
              tags$div(class = "sci-card",
                       tags$div(class = "sci-card-header",
                                tags$div(class = "card-icon icon-amber", tags$i(class = "fa fa-eye")),
                                tags$h4("Sheet Preview")
                       ),
                       tags$div(class = "sci-card-body",
                                selectInput("sheetSelect", label = "Select worksheet:", choices = NULL),
                                DT::dataTableOutput("excelPreview")
                       )
              ),

              # Footer
              tags$div(class = "app-footer",
                       "Developed by ", tags$a("Philip Duah", href="mailto:dpbxc@mst.edu"),
                       "  ·  Missouri University of Science & Technology",
                       "  ·  NSF Award No. 2219086"
              )
      ),

      # ── ABOUT TAB ────────────────────────────────────────────────────────
      tabItem(tabName = "about",

              tags$div(class = "page-hero",
                       tags$div(class = "hero-badge", tags$i(class="fa fa-flask", style="margin-right:5px;"), "Open Science"),
                       tags$h2("About This Application"),
                       tags$p("Background, acknowledgments, and the vision behind the MSMG SimaPro CSV Optimizer.")
              ),

              tags$div(class = "sci-card",
                       tags$div(class = "sci-card-header",
                                tags$div(class = "card-icon icon-green", tags$i(class="fa fa-info-circle")),
                                tags$h4("Overview")
                       ),
                       tags$div(class = "sci-card-body",
                                tags$div(class = "content-section",
                                         tags$h3("MSMG SimaPro CSV Optimizer"),
                                         tags$p("This application was developed and deployed by ", tags$strong("Philip Duah"),
                                                " using the R programming language as part of the project ",
                                                tags$em("ECO‑CBET: GOALI: CAS‑Climate: Expediting Decarbonization of the
                     Cement Industry through Integration of CO₂ Capture and Conversion"),
                                                " (NSF Project No. 2219086)."),

                                         tags$h4("Acknowledgments"),
                                         tags$p("I extend my gratitude to ", tags$strong("Professor Massimo Pizzol"),
                                                ", whose earlier work in Python inspired the creation of this
                     user-friendly web application aimed at advancing life cycle assessment (LCA) modeling."),
                                         tags$ul(
                                           tags$li(tags$a("Professor Pizzol's LCA Tutorials",
                                                          href = "https://moutreach.science/2022/08/15/teaching-videos.html",
                                                          target = "_blank")),
                                           tags$li(tags$a("Professor Pizzol's SimaPro CSV Converter (GitHub)",
                                                          href = "https://github.com/massimopizzol/Simapro-CSV-converter",
                                                          target = "_blank"))
                                         ),

                                         tags$h4("Roadmap & Vision"),
                                         tags$p("This initial version reflects a broader vision of enhancing interoperability and
                     accessibility within LCA and Engineering. Planned updates include:"),
                                         tags$ul(
                                           tags$li("Extended support for additional LCI data formats"),
                                           tags$li("Direct integration with ecoinvent and GLAD nomenclature mappings"),
                                           tags$li("Batch validation and data quality scoring"),
                                           tags$li("API access for automated pipeline workflows")
                                         ),

                                         tags$div(class = "citation-box",
                                                  tags$strong("Contact & Correspondence:"),
                                                  tags$br(),
                                                  "Questions and inquiries should be directed to: ",
                                                  tags$a("dpbxc@mst.edu", href = "mailto:dpbxc@mst.edu")
                                         )
                                )
                       )
              )
      ),

      # ── INSTRUCTIONS TAB ─────────────────────────────────────────────────
      tabItem(tabName = "instructions",

              tags$div(class = "page-hero",
                       tags$div(class = "hero-badge", tags$i(class="fa fa-book-open", style="margin-right:5px;"), "User Guide"),
                       tags$h2("How to Use the Optimizer"),
                       tags$p("A step-by-step guide from Excel preparation to SimaPro import.")
              ),

              tags$div(class = "sci-card",
                       tags$div(class = "sci-card-header",
                                tags$div(class = "card-icon icon-green", tags$i(class="fa fa-list-ol")),
                                tags$h4("Conversion Workflow")
                       ),
                       tags$div(class = "sci-card-body",
                                tags$div(class = "content-section",
                                         tags$ol(class = "step-list",
                                                 tags$li("Prepare your life cycle inventory in Excel following the required format.
                         Navigate to the ", tags$strong("Converter"), " tab and upload your
                         ", tags$code(".xlsx"), " or ", tags$code(".xls"), " file."),
                                                 tags$li("Click ", tags$strong("Convert to SimaPro CSV"), " to process all worksheets.
                         The system will generate one CSV per worksheet and package them into a ZIP archive."),
                                                 tags$li("Download the ZIP archive using the ", tags$strong("Download ZIP"), " button.
                         Extract all files to a folder of your choice."),
                                                 tags$li("Open SimaPro and navigate to ", tags$strong("File → Import"),
                                                         " to begin the import process."),
                                                 tags$li(
                                                   "Configure the import settings as follows:",
                                                   tags$ul(
                                                     tags$li("Set the file format to ", tags$strong("SimaPro CSV")),
                                                     tags$li("Click ", tags$strong("Add"), " and navigate to your converted CSV file"),
                                                     tags$li("Leave the ", tags$strong("Mapping file"), " section empty"),
                                                     tags$li("For object link method, choose ",
                                                             tags$strong("Try to link imported objects to existing objects first")),
                                                     tags$li("Set CSV format separator to ", tags$strong("Semicolon"), " (Tab also works)"),
                                                     tags$li("Enable ", tags$strong("Replace existing processes with equal identifiers"),
                                                             " under Other Options")
                                                   )
                                                 )
                                         )
                                )
                       )
              ),

              tags$div(class = "sci-card",
                       tags$div(class = "sci-card-header",
                                tags$div(class = "card-icon icon-amber", tags$i(class="fa fa-exchange-alt")),
                                tags$h4("Supported Exchange Type Codes")
                       ),
                       tags$div(class = "sci-card-body",
                                tags$div(class = "content-section",
                                         tags$p("Use the following codes in column A of your Excel worksheet to classify flows:"),
                                         tags$div(
                                           style = "display: grid; grid-template-columns: repeat(auto-fill, minmax(200px,1fr)); gap: 10px; margin-top: 12px;",
                                           lapply(list(
                                             list(code="(blank)",         label="Materials / Fuels",   color="#1a6b4a"),
                                             list(code="Raw",             label="Resources",           color="#1a6b9e"),
                                             list(code="Air",             label="Emissions to Air",    color="#5a6a7a"),
                                             list(code="Water",           label="Emissions to Water",  color="#1a7a9e"),
                                             list(code="Soil",            label="Emissions to Soil",   color="#7a5a2a"),
                                             list(code="Waste",           label="Final Waste Flows",   color="#c0392b"),
                                             list(code="Social",          label="Social Issues",       color="#6a3a9e"),
                                             list(code="Economic",        label="Economic Issues",     color="#c8922a"),
                                             list(code="Wastetotreatment",label="Waste to Treatment",  color="#2a7a4a")
                                           ), function(x) {
                                             tags$div(
                                               style = sprintf("background: var(--white); border: 1.5px solid var(--border);
                                     border-left: 4px solid %s; border-radius: 8px;
                                     padding: 10px 14px;", x$color),
                                               tags$div(style = sprintf("font-family: monospace; font-size: 13px; font-weight: 700; color: %s;", x$color), x$code),
                                               tags$div(style = "font-size: 12px; color: var(--ink-muted); margin-top: 3px;", x$label)
                                             )
                                           })
                                         )
                                )
                       )
              )
      )
    )
  )
)

# ─────────────────────────────────────────────────────────────────────────────
# SERVER
# ─────────────────────────────────────────────────────────────────────────────
server <- function(input, output, session) {
  values <- reactiveValues(
    sheets = NULL,
    csv_files = list(),
    processing_complete = FALSE,
    processing_log = "No file processed yet."
  )

  observeEvent(input$file, {
    req(input$file)
    tryCatch({
      sheet_names <- excel_sheets(input$file$datapath)
      values$sheets <- sheet_names
      updateSelectInput(session, "sheetSelect", choices = sheet_names)
      if(length(sheet_names) > 0)
        updateSelectInput(session, "sheetSelect", selected = sheet_names[1])
      values$processing_log <- paste("Detected", length(sheet_names), "sheet(s) in the uploaded file.")
      # Update stat card
      session$sendCustomMessage("updateStat", list(id = "stat-sheets", val = length(sheet_names)))
    }, error = function(e) {
      values$processing_log <- paste("Error reading Excel file:", e$message)
    })
  })

  output$excelPreview <- DT::renderDataTable({
    req(input$file, input$sheetSelect)
    tryCatch({
      df <- suppressWarnings(read_excel(input$file$datapath,
                                        sheet = input$sheetSelect,
                                        col_names = FALSE,
                                        .name_repair = "minimal"))
      DT::datatable(df,
                    options = list(
                      scrollX    = TRUE,
                      pageLength = 15,
                      lengthMenu = c(5, 10, 15, 25, 50),
                      dom        = 'lfrtip',
                      language   = list(search = "Filter:")
                    ),
                    class = "stripe hover",
                    rownames = FALSE
      )
    }, error = function(e) {
      values$processing_log <- paste(values$processing_log, "\nError previewing sheet:", e$message)
      data.frame(Error = paste("Could not read sheet:", e$message))
    })
  })

  observeEvent(input$convert, {
    req(input$file)
    loading_html <- list(
      spin_flower(),
      tags$div(style = "margin-top: 20px; font-family: 'DM Serif Display', serif; font-size: 20px; color: #ffffff;", "Converting…"),
      tags$div(style = "margin-top: 8px; font-size: 13px; color: rgba(255,255,255,0.65);", "Please wait while your LCI data is processed")
    )
    w <- Waiter$new(html = loading_html, color = "rgba(15, 25, 35, 0.88)")
    w$show()
    Sys.sleep(2)
    values$processing_log <- "Starting conversion process...\n"
    values$csv_files <- list()
    values$processing_complete <- FALSE
    temp_dir <- file.path(tempdir(), paste0("simapro_", format(Sys.time(), "%Y%m%d_%H%M%S")))
    dir.create(temp_dir, showWarnings = FALSE, recursive = TRUE)
    values$processing_log <- paste0(values$processing_log, "Temporary directory created.\n")
    for (sheet_name in values$sheets) {
      values$processing_log <- paste0(values$processing_log, "Processing sheet: ", sheet_name, "\n")
      tryCatch({
        df <- suppressWarnings(read_excel(input$file$datapath,
                                          sheet      = sheet_name,
                                          col_names  = FALSE,
                                          .name_repair = "minimal",
                                          guess_max  = 10000))
        values$processing_log <- paste0(values$processing_log,
                                        "  — ", nrow(df), " rows × ", ncol(df), " columns\n")
        safe_name  <- gsub("[^a-zA-Z0-9_.-]", "_", sheet_name)
        output_file <- file.path(temp_dir, paste0(safe_name, ".csv"))
        result <- to_spcsv(df, output_file)
        if (result && file.exists(output_file)) {
          file_size <- file.info(output_file)$size
          if (file_size > 100) {
            values$csv_files[[sheet_name]] <- output_file
            values$processing_log <- paste0(values$processing_log, "  ✓ Success (", file_size, " bytes)\n")
          } else {
            values$processing_log <- paste0(values$processing_log, "  ⚠ Warning: file may be incomplete\n")
          }
        } else {
          values$processing_log <- paste0(values$processing_log, "  ✗ Error: file creation failed\n")
        }
      }, error = function(e) {
        values$processing_log <- paste0(values$processing_log, "  ✗ Error: ", e$message, "\n")
      })
    }
    n_csv <- length(values$csv_files)
    if (n_csv > 0) {
      values$processing_complete <- TRUE
      values$processing_log <- paste0(values$processing_log,
                                      "\n✓ Conversion complete — ", n_csv, " CSV file(s) ready for download.\n")
      session$sendCustomMessage("updateStat", list(id = "stat-csvs", val = n_csv))
    } else {
      values$processing_log <- paste0(values$processing_log,
                                      "\n✗ Conversion failed. No CSV files were created.\n  Please verify that your Excel file matches the required format.\n")
    }
    w$hide()
  })

  output$processingStatus <- renderText({ values$processing_log })

  output$downloadZip <- downloadHandler(
    filename = function() { paste0("SimaPro_CSV_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".zip") },
    content  = function(file) {
      if (!values$processing_complete || length(values$csv_files) == 0) {
        showModal(modalDialog(
          title     = "No Files Available",
          "Please convert your Excel file first before downloading.",
          easyClose = TRUE,
          footer    = modalButton("OK")
        ))
        file.create(file); return()
      }
      files_to_zip <- unlist(values$csv_files)
      zip_dir      <- unique(dirname(files_to_zip))[1]
      current_dir  <- getwd()
      setwd(zip_dir)
      on.exit(setwd(current_dir), add = TRUE)
      tryCatch({
        zip::zip(zipfile = file, files = basename(files_to_zip), mode = "cherry-pick")
      }, error = function(e) {
        values$processing_log <- paste0(values$processing_log, "\nError creating ZIP: ", e$message)
        file.create(file)
      })
    },
    contentType = "application/zip"
  )

  # JS handler to update stat cards
  session$onFlushed(function() {
    shinyjs_code <- "
    Shiny.addCustomMessageHandler('updateStat', function(msg) {
      var el = document.getElementById(msg.id);
      if (el) { el.textContent = msg.val; }
    });
    "
    session$sendCustomMessage("evaljs", shinyjs_code)
  }, once = TRUE)
}

shinyApp(ui, server)
