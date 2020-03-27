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

#' Create pdf randomization lists from a list of blockrand
#' generated data.frame 
#'
#' Create pdf randomization lists from a list of blockrand
#' generated data.frame
#' 
#' @param x a single data.frame or a named randlist (that is a
#'     data.frame with id and treatment columns). Names are used for
#'     file naming
#' @param path_prefix path prefix of the files to save
#'     in (overwriting the contents).
#' @param footer a character vector used as page central
#'     footer(s). Must be of length 1 if x is a data.frame or of the
#'     same length of x, if it's a list.
#' @export
randlist2pdf <- function(x = NULL,
                         path_prefix = '/tmp/randlist',
                         footer = "") {

    ## make a list of data.frames
    x <- normalize_randlists(x)
    xnames <- lbmisc::preprocess_varnames(names(x), dump_rev = FALSE)
    ## check that these are randlists
    are_rl <- lapply(x, function(x) all(c('id', 'treatment') %in% names(x)))
    if (!all(unlist(are_rl)))
        stop('x has not id and/or treatment variable/s')

    if (!(is.character(footer) && length(footer) %in% c(1L, length(x)))){
        msg <- c("footer must be a character of length 1 ",
                 "or of the same number of x's data.frames")
        stop(msg)
    }
    
    if (!(is.character(path_prefix) && length(path_prefix) == 1L))
        stop('path_prefix must be a character of length 1')

    ## modify each data frame to a proper output format
    x <- lapply(x, function(rl){
        ## Add needed columns
        new_vars <- c("cognome_pz", "nome_pz", "cognome_dr", "nome_dr",
                      "ora", "data", "sigla", "note")
        rl[new_vars] <- NA
        ## Keep only what's needed
        needed_vars <- c("id", "cognome_pz",  "nome_pz",
                         "treatment", "cognome_dr", "nome_dr",
                         "ora", "data", "sigla", "note" )
        
        rl <- rl[needed_vars]
        ## change to alfanumeric id
        if (is.numeric(rl$id))
            rl$id <- lbmisc::to_00_char(rl$id, floor(log10(max(rl$id))) + 1)

        return(rl)
    })
    
    ## occorre aggiungere il nome dello strato
    files <- paste(path_prefix, xnames, sep = '_')
    Map(function(db, footer, file){
        make_randlist_pdf(x = db, cfoot = footer, f = file)
    }, x, as.list(footer), as.list(files))
    invisible(NULL)
}


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



#' Create pdf, envelopes and xlsx randomization lists for a
#' stratified, blocked study
#'
#' @param pi global PI
#' @param acronym study acronym
#' @param sample_size total sample size (a randomization list with
#'     this numerosity will be created for each stratum/element of
#'     stratas)
#' @param seed random seed
#' @param treatment_levels labels used to identify the groups
#' @param block_size blocks dimensions
#' @param centres id of involved centres
#' @param stratas label for each strata
#' @param local_pis a string for footers of the printed lists, in the
#'     same number of stratas
#' @param testing if TRUE it will generate a list using a different
#'     seed, for testing purposese (eg EDC setup)
#' @param print_checks print performed checks on the randomization
#'     lists
#' @param export format of lists exporting (can have pdf, xlsx and/or
#'     envelopes)
#' @param env_params envelopes parameters for printing (passed to
#'     blockrand::plotblockrand)
#' @export
blocked_stratified_randlist <- 
    function(pi = '',                   # cognome del pi globale
             acronym = '',            # acronimo dello studio
             sample_size = NA,               # dimensione campione
             seed = NA,                   # seme casuale
             treatment_levels = c("C","T"),  # etichette gruppi
             block_size = c(2L, 4L, 6L),     # dimensione blocchi blocking
             centres = c('asmn'),            # centri
             stratas = centres,              
             local_pis = c('PI Cognome Nome (ASMN/AUSL)'),   # pi per centri
             testing = FALSE,                # test or official randlist
             print_checks = TRUE,
             export = c('pdf', 'xlsx', 'envelopes'),
             env_params = list(width = 11, height = 8))
{
    
    if (length(stratas) != length(local_pis))
        stop("stratas and local_pis must have the same length")
    
    ## altri parametri utili generati automaticamente
    names(treatment_levels) <- treatment_levels
    footers <- sprintf("Studio %s - %s - [Strato: %s]", 
                       acronym, local_pis, stratas)
    mono_multicentrico <- if (length(centres) > 1) 'multicentrico' 
                          else 'monocentrico'
    n_gruppi_trattamento <- length(treatment_levels)
    dimensione_blocchi <- paste(block_size, collapse = ', ')

    ## generazione della lista
    used_seed <- if (testing) 12345 else seed
    set.seed(used_seed)
    randlists <- lapply(stratas, function(x) {
        blockrand::blockrand(n = sample_size,
                             num.levels = length(treatment_levels),
                             stratum = x, 
                             levels = treatment_levels,
                             block.sizes = rep(block_size / 2L, 2L))
    })
    names(randlists) <- stratas
    ## Aggiunta id di strato
    ## OLD letters
    ## strata_prefix <- c(paste0(LETTERS[seq_along(names(randlists))], '-'))
    ## NEW numeric (000-000)
    strata_prefix <-
        paste0(lbmisc::to_00_char(seq_along(names(randlists)), 3L), '-')

    add_strata_prefix <- function(rl, prefix) {
        nmax_digits <- ceiling(max(log10(rl$id)))
        rl$id <- paste0(prefix, lbmisc::to_00_char(rl$id, nmax_digits))
        rl
    }
    randlists <- Map(f = add_strata_prefix, randlists, strata_prefix)

    ## ------
    ## CHECKS
    ## ------
    if (print_checks){
        f <- function(x, fun, as.vec = TRUE){
            rval <- lapply(X = x, FUN = fun)
            if (as.vec) unlist(rval) else rval
        }
        message('Numerosità per strato')
        print(f(randlists, nrow))
        message('Numero di blocchi per strato')
        print(f(randlists,  function(x) length(unique(x$block.id))))
        message('Bilanciamento complessivo per strato')
        print(do.call('rbind', f(randlists, function(x) table(x$treatment), FALSE)))
        message('Bilanciamento entro ciascun blocco dello strato')
        print(f(randlists, function(x) {    
            tmp <- as.matrix(table(x$block.id, x$treatment))
            all(tmp[, 1] == tmp[, 2])
        }))
    }

    ## ------
    ## OUTPUT
    ## ------
    testing_string <- if (testing) 'TESTING' else 'OFFICIAL'
    if ('pdf' %in% export){
        pdf_path  <- sprintf("/tmp/%s_%s_%s_pdf_randomization_lists",
                             pi, acronym, testing_string)
        lbrct::randlist2pdf(x = randlists, 
                            path_prefix = pdf_path, 
                            footer = footers)
    }
    
    if ('xlsx' %in% export){
        xlsx_path <- sprintf("/tmp/%s_%s_%s_raw_randomization_lists",
                             pi, acronym, testing_string)
        select <- lapply(randlists, 
                         function(x) x[,c('id', 'stratum', 'treatment')])
        openxlsx::write.xlsx(x = select, file = paste0(xlsx_path, '.xlsx'))
    }

    if ('envelopes' %in% export){
        pdf_path  <- sprintf("/tmp/%s_%s_%s_envelopes",
                             pi, acronym, testing_string)
        lbrct::randlist2envelopes(x = randlists,
                                  path_prefix = pdf_path,
                                  study_acronym = acronym,
                                  env_params = env_params)
    }
    
    randlists
}


#' Create randomization envelopes from a list of blockrand
#' generated data.frame 
#'
#' Create randomization envelopes from a list of blockrand
#' generated data.frame
#' 
#' @param x a single data.frame or a named randlist (that is a
#'     data.frame with id and treatment columns). Names are used for
#'     file naming
#' @param path_prefix path prefix of the files to save
#'     in (overwriting the contents).
#' @param study_acronym a character vector used as acronym
#' @param env_params list of parameters passed to blockrand::plotblockrand
#' @export
randlist2envelopes <- function(x = NULL,
                               path_prefix = '/tmp/envelopes',
                               study_acronym = "",
                               env_params = list()) {

    ## make a list of data.frames
    x <- normalize_randlists(x)
    xnames <- lbmisc::preprocess_varnames(names(x), dump_rev = FALSE)
    ## check that these are randlists
    are_rl <- lapply(x, function(x) all(c('id', 'treatment') %in% names(x)))
    if (!all(unlist(are_rl)))
        stop('x has not id and/or treatment variable/s')

    if (!(is.character(path_prefix) && length(path_prefix) == 1L))
        stop('path_prefix must be a character of length 1')
    
    ## occorre aggiungere il nome dello strato
    files <- paste0(path_prefix, '_', xnames, '.pdf')
    Map(function(x, f){
        study_acronym_label <- sprintf("Study: %s", study_acronym)
        blockrand_text <- list(
            top = list(text = c(study_acronym_label,
                                "Strata: %STRAT%",
                                'Patient ID: %ID%',
                                'Treatment: %TREAT%'),
                       font = c(1,1,1,2)),
            middle = list(text = c(study_acronym_label,
                                   "Strata: %STRAT%",
                                   "Patient ID: %ID%"),
                          font = 1),
            bottom = "")
        plot_params <- c(list(x = x, file = f,
                              blockrand.text = blockrand_text,
                              cut.marks = TRUE),
                         env_params)
        do.call(blockrand::plotblockrand, plot_params)
    }, x, as.list(files))
    invisible(NULL)
}

