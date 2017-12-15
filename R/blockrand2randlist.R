## normalize_x in order to handle both single data.frame and list
## of data.frames
normalize_x <- function(x){
    if (is.data.frame(x)){
        x <- list(x)
        names(x) <- '1'
    } else if (is.list(x)) {
        if (is.null(names(x)))
            names(x) <- as.character(seq_len(length(x)))
    } else
        stop('x must be a data.frame or a list of data frames')

    x
}


#' Create a randomization list from a list of blockrand generated data.frame 
#'
#' Create a randomization list from a list of blockrand generated data.frame
#' 
#' @param x named list of blockrand data.frame (names will be used as
#'     sheet name)
#' @param f path to file to save in (overwriting the contents). If
#'     NULL the list is displayed.
#' @param footer a character vector used as page central
#'     footer(s). Page numbers are on the left, while date/time is on
#'     the right.
#' @export
blockrand2randlist <- function(x,
                               f = '/tmp/randlist',
                               footer = "") {

    x <- normalize_x(x)
    sheet_names <- names(x)

    if (!(is.character(footer) && length(footer) %in% c(1L, length(x)))){
        msg <- c("footer must be a character of length 1 ",
                 "or of the same number of x's data.frames")
        stop(msg)
    }
    
    if (!((is.character(f) && length(f) == 1L) || is.null(f)))
        stop('f must be a character of length 1 or NULL')

    ## actual n for each strata (depending on 
    actual_n <- unlist(lapply(x, nrow))
    
    ## modify each data frame to a proper output format
    x <- lapply(x, function(rl){
        ## Add needed columns
        rl[c("Cognome.pz", "Nome.pz",
             "Cognome.dr", "Nome.dr",
             "Ora", "Data", "Sigla", "Note"
             )] <- NA
        ## Remove unneeded stuff
        rl <- rl[c("id", "Cognome.pz",  "Nome.pz",
                   "treatment", "Cognome.dr", "Nome.dr",
                   "Ora", "Data", "Sigla", "Note" )]
        ## Rename columns
        names(rl) <- c("ID", "Cognome",  "Nome",
                       "TRAT", "Cognome", "Nome",
                       "Ora", "Data", "Sigla di chi risponde", 
                       "Note")
        ## change to alfanumeric id
        if (is.numeric(rl$ID))
            rl$ID <- lbmisc::to_00_char(rl$ID, floor(log10(max(rl$ID))) + 1)

        return(rl)
    })

    ## -----------------
    ## LATEX/PDF OUTPUT
    ## -----------------

    ## occorre aggiungere il nome dello strato
    files <- paste(f, preprocess_varnames(names(x)), sep = '_')
    Map(function(db, footer, file){
        make_pdf_randlist(db, cfoot = footer, f = file)
    }, x, as.list(footers), as.list(files))
                  
    ## -----------------
    ## EXCEL STUFF BELOW
    ## -----------------
    
    ## Sheets' header
    header_inv <- header_tc <- matrix(c("Dati del paziente",
                                        rep(NA,3),
                                        "Dati di chi chiama",
                                        NA,
                                        "Dati della chiamata",
                                        NA,
                                        "Sigla di chi risponde",
                                        "Note"
                                        ), nrow = 1)
    header_inv[1, c(5,9)] <- c("Dati di chi risponde", "Sigla di chi chiama")

    ## Setup the workbook
    wb <- openxlsx::createWorkbook()
    worksheet_creator <- function(s, foot) {
        openxlsx::addWorksheet(wb = wb, sheetName = s,
                               footer = c("Page &[Page] of &[Pages]", # left
                                          foot,                       # centre
                                          "&[Date] &[Time]") )        # right
    }
    Map(worksheet_creator, as.list(sheet_names), as.list(footer))
  
    ## Page setup variables
    ColConversionFactor <- 6
    RowConversionFactor <- 5.5^2
    headerColWidths <- c(2.2, 3.1, 3.1, 1.5, 3, 3, 1.8, 2.5, 2.2, 4.2)#cm
    headerRowHeights <- 1 # cm
    otherRowHeights <- 2.3# cm
    margins <- 0.4 # inches == 1 cm
    
    row_heights <- lapply(actual_n, function(n) {
        ## rep(c(headerRowHeights, otherRowHeights), c(2, nrow(x[[1]])))
        rep(c(headerRowHeights, otherRowHeights), c(2, n))
    })

    rlStyle <- openxlsx::createStyle(fontName = "Arial", 
                                     fontSize = 12,
                                     border = "TopBottomLeftRight",
                                     textDecoration = "bold",
                                     halign = "center",
                                     valign = "center",
                                     wrapText = TRUE)

    ## Setup each sheet/dataset
    Map(sheet_names,
        row_heights,
        f = function(s, rh){
               
            openxlsx::pageSetup(wb = wb,
                                sheet = s,
                                scale = 83, # to make it all fits
                                orientation = "landscape",
                                fitToWidth = TRUE, 
                                left = margins,
                                right = margins,
                                top = margins,
                                bottom = margins * 1.5,
                                printTitleRows = 1:2)
            
            openxlsx::setColWidths(wb = wb,
                                   sheet = s,
                                   cols = 1:10,
                                   widths = headerColWidths * ColConversionFactor)
            
            openxlsx::setRowHeights(wb = wb,
                                    sheet = s, 
                                    rows = seq_len(length(rh)),
                                    heights = rh * RowConversionFactor)
            
            openxlsx::addStyle(wb = wb,
                               sheet = s,
                               style = rlStyle, 
                               cols = seq_len(ncol(x[[s]])),
                               rows = seq_len(nrow(x[[s]]) + 2), # +2 per l'header 
                               gridExpand = TRUE)
    
            ## Merge Cells dell'header
            openxlsx::mergeCells(wb = wb, sheet = s, cols = 1:4, rows = 1)
            openxlsx::mergeCells(wb = wb, sheet = s, cols = 5:6, rows = 1)
            openxlsx::mergeCells(wb = wb, sheet = s, cols = 7:8, rows = 1)
            openxlsx::mergeCells(wb = wb, sheet = s, cols =   9, rows = 1:2)
            openxlsx::mergeCells(wb = wb, sheet = s, cols =  10, rows = 1:2)
        })


    ## -----------------------------
    ## Trial Center List - full list
    ## -----------------------------
    wb_tc <- wb
    lapply(sheet_names, function(s) {
        ## header
        openxlsx::writeData(wb = wb,
                            sheet = s,
                            x = header_tc,
                            colNames = FALSE)
        ## data
        openxlsx::writeData(wb = wb_tc,
                            sheet = s,
                            x = x[[s]],
                            startRow = 2)
    })
    if (is.null(f)) {
        openxlsx::openXL(wb_tc)
    } else {
        tc_file <- paste0(f, '_TRIAL_CENTER.xlsx')
        openxlsx::saveWorkbook(wb = wb_tc, file = tc_file, overwrite = TRUE)
    }

    ## -----------------------------
    ## Investigators - blanked list
    ## -----------------------------
    if (!is.null(f)) {
        wb_inv <- wb
        lapply(sheet_names, function(s) {
            ## header
            openxlsx::writeData(wb = wb,
                                sheet = s,
                                x = header_inv,
                                colNames = FALSE)
            ## data
            tmp <- x[[s]]
            tmp$TRAT <- NA
            openxlsx::writeData(wb = wb_inv,
                                sheet = s,
                                x = tmp,
                                startRow = 2)
        })
        
        inv_file <- paste0(f, '_INVESTIGATORS.xlsx')
        openxlsx::saveWorkbook(wb = wb_inv, file = inv_file, overwrite = TRUE)

    }
}





    randlist_header <- "\\documentclass[a4paper, 12pt]{article}
\\renewcommand*\\familydefault{\\sfdefault}
\\usepackage[T1]{fontenc}
\\usepackage[utf8]{inputenc}
\\usepackage[english, italian]{babel}
\\usepackage[yyyymmdd]{datetime}
\\renewcommand{\\dateseparator}{-}
\\usepackage[
landscape,
a4paper,
top=0.5cm,
bottom=1.5cm,
left=0.3cm,
right=0.3cm
]{geometry}
\\usepackage{lastpage}
\\usepackage{multirow, makecell}
\\usepackage{fancyhdr}
\\pagestyle{fancy}
\\renewcommand{\\headrulewidth}{0pt}
\\lfoot{Page \\thepage{} of \\pageref{LastPage}}
\\rfoot{\\today{} \\currenttime}
\\cfoot{%s}
\\usepackage{longtable,array}
\\newcolumntype{G}{>{\\centering\\arraybackslash}p{3cm}}
\\newcolumntype{M}{>{\\centering\\arraybackslash}p{1.8cm}}
\\newcolumntype{P}{>{\\centering\\arraybackslash}p{1.5cm}}
\\begin{document}
\\renewcommand{\\arraystretch}{2.5}
\\begin{longtable}{|P|G|G|P|G|G|P|G|M|G|}
  \\hline
  \\multicolumn{4}{|c|}{\\textbf{Dati del paziente}} &
  \\multicolumn{2}{c|}{\\textbf{Dati di chi %s}} &
  \\multicolumn{2}{c|}{\\textbf{Dati della chiamata}} & 
  \\multirowcell{2}{\\textbf{Sigla} \\\\ \\textbf{di chi} \\\\ \\textbf{%s}}&
  \\multirowcell{2}{\\textbf{Note}} \\\\
  \\cline{1-8}
  \\textbf{ID} & 
  \\textbf{Cognome} & 
  \\textbf{Nome} & 
  \\textbf{TRAT} & 
  \\textbf{Cognome} & 
  \\textbf{Nome} & 
  \\textbf{Ora} & 
  \\textbf{Data} &  &  \\\\ 
  \\hline
  \\endhead
"
    footer <- "\\end{longtable}
\\end{document}
"


make_pdf_randlist <- function(x = NULL,
                              cfoot = NULL,
                              f = 'randomization_list',
                              compile = TRUE,
                              view = FALSE)
{

    target <- list('TRIAL_CENTER', 'INVESTIGATOR')
    texfiles <- lapply(target, function(y) sprintf("%s_%s.tex", f, y))
    pdffiles <- lapply(target, function(y) sprintf("%s_%s.pdf", f, y))
    dbs <- list('tc' = x, 'inv' = {x$TRAT <- NA; x})
    
    cr_base <- c('chiama', 'risponde')
    cr <- list('tc' = cr_base, 'inv' = rev(cr_base))
    
    make_header <- function(cr) sprintf(randlist_header, cfoot, cr[1], cr[2])
    headers <- Map(make_header, cr)
    
    make_table_contents <- function(x){    
        xtable::print.xtable(xtable::xtable(x),
                             print.results = FALSE,
                             comment = FALSE,
                             include.colnames = FALSE, 
                             include.rownames = FALSE,
                             only.contents = TRUE,
                             hline.after = seq_len(nrow(x)))
    }
    
    table_contents <- lapply(dbs, make_table_contents)
    
    tex_maker <- function(h, tc, foot, of) cat(h, tc, foot, file = of)
    
    Map(tex_maker, headers, table_contents, list(footer), texfiles)

    if (compile){
        browser()
        compiler <- function(f) {
            oldpwd <- getwd()
            on.exit(setwd(oldpwd))
            setwd(dirname(path.expand(f)))
            tools::texi2pdf(f, clean = TRUE)
        }
        lapply(texfiles, compiler)
    }
    
    if (view){
        viewer <- function(x) system(sprintf('evince %s &', x))
        lapply(pdffiles, viewer)
    }
}


## make_pdf_randlist(x = db)
