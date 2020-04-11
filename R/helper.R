## normalize_randlists in order to handle both single data.frame and list
## of data.frames
normalize_randlists <- function(x){
    if (is.null(x)){
        stop("x can't be null")
    } else if (is.data.frame(x)){
        x <- list(x)
        names(x) <- '1'
    } else if (is.list(x)) {
        if (is.null(names(x)))
            names(x) <- as.character(seq_len(length(x)))
    } else
        stop('x must be a data.frame or a list of data frames')

    x
}

## function to create a full xlsx structure (eg for gdocs)
format_xlsx <- function(x, local_pi){
    ## x is a single data.frame
    ## local.pi is the pi for the centre
    ## add local pi to stratum
    x$stratum <- sprintf("%s (%s)", x$stratum, local_pi)
    x$cognome_pz         <- NA
    x$nome_pz            <- NA
    x$cognome_chi_chiama <- NA
    x$nome_chi_chiama    <- NA
    x$data_chiamata      <- NA
    x$ora_chiamata       <- NA
    x$chi_randomizza     <- NA
    x$note               <- NA
    var_order <- c('stratum', 'id', 'cognome_pz', 'nome_pz',
                   'treatment', 'cognome_chi_chiama',
                   'nome_chi_chiama', 'ora_chiamata',
                   'data_chiamata', 'chi_randomizza', 'note')
    x[, var_order]
}

## exporta single xlsx randomization list
xlsx_exporter <- function(x, xn, do_export){
    xn <- gsub(' ', '_', tolower(xn))
    ## filepath <- sprintf('/tmp/%s_%s.xlsx', acronym, xn)
    filepath <- sprintf('/tmp/%s.xlsx', xn)
    if (do_export) openxlsx::write.xlsx(x, file = filepath)
    else NULL
}

## worker for lists export
make_randlist_pdf <- function(x = NULL, cfoot = NULL, f = NULL){

    target <- list('TRIAL_CENTER', 'INVESTIGATORS')
    texfiles <- lapply(target, function(y) sprintf("%s_%s.tex", f, y))
    pdffiles <- lapply(target, function(y) sprintf("%s_%s.pdf", f, y))
    dbs <- list('tc' = x, 'inv' = {x$treatment <- NA; x})

    ## composizione documento
    cr_base <- c('chiama', 'risponde')
    crs <- list('tc' = cr_base, 'inv' = rev(cr_base))
    make_header <- function(x) sprintf(randlist_header, cfoot, x[1], x[2])
    headers <- Map(make_header, crs)
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
    randlist_footer <- list("\\end{longtable}\\end{document}")
    tex_maker <- function(h, tc, foot, of) cat(h, tc, foot, file = of)
    Map(tex_maker, headers, table_contents, randlist_footer, texfiles)

    ## compiling
    compiler <- function(f) {
        oldpwd <- getwd()
        on.exit(setwd(oldpwd))
        setwd(dirname(path.expand(f)))
        tools::texi2pdf(f, clean = TRUE)
    }
    lapply(texfiles, compiler)
    invisible(NULL)
}

## header delle liste di randomizzazione
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
