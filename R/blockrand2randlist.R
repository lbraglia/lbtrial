#' Create randomization lists
#'
#' Create pdf, envelopes and xlsx randomization lists for a
#' stratified, blocked study and export them
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
#' @param export logical (recicled, by default TRUE), which
#'     randomization lists to export
#' @param export_format format of lists exporting (can have pdf, xlsx
#'     and/or envelopes)
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
             centres = c('auslre'),            # centri
             stratas = centres,              
             local_pis = c('PI Cognome Nome (AUSL RE-Irccs)'), # pi per centri
             testing = FALSE,                # test or official randlist
             print_checks = TRUE,
             export = TRUE,
             export_format = c('pdf', 'xlsx', 'xlsx_minimal', 'envelopes'),
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

    ## aggiungere anche i pi locali male non fa per le esportazioni
    names(randlists) <-
        lbmisc::preprocess_varnames(paste(stratas, local_pis),
                                    dump_rev = FALSE)

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
    ## export if there's at least a centre to be exported
    if (any(export)) {
    
        std_path <- sprintf("/tmp/%s%s", acronym,
                            if (testing) '_TESTING' else '')
        if ('pdf' %in% export_format){
            pdf_path  <- paste0(std_path, "_randlist")
            ## pdf_path <- std_path
            lbrct::randlist2pdf(x = randlists, 
                                path_prefix = pdf_path, 
                                footer = footers,
                                export = export)
        }
        
        if ('xlsx' %in% export_format){
            
            lbrct::randlist2xlsx(x = randlists,
                                 local_pis = local_pis, 
                                 export = export)

            
           
        }

        if ('xlsx_minimal' %in% export_format){
            xlsx_path <- std_path
            selector <- function(x) x[,c('id', 'stratum', 'treatment')]
            select <- lapply(randlists, selector)
            select <- select[export]
            openxlsx::write.xlsx(x = select, file = paste0(xlsx_path, '.xlsx'))
        }
        
        if ('envelopes' %in% export_format){
            pdf_path  <- paste0(std_path, "_envelopes")
            lbrct::randlist2envelopes(x = randlists,
                                      path_prefix = pdf_path,
                                      study_acronym = acronym,
                                      env_params = env_params,
                                      export = export
                                      )
        }
    }
    
    randlists
}




#' Create xlsx randomization lists from a list of blockrand
#' generated data.frame 
#'
#' Create xlsx randomization lists from a list of blockrand
#' generated data.frame
#' 
#' @param x a single data.frame or a named randlist (that is a
#'     data.frame with id and treatment columns). Names are used for
#'     file naming
#' @param local_pi local principal investigators
#' @param export logical recycled, wheter to export or not a list
#'     (used to export selectively)
#' @export
randlist2xlsx <- function(x = NULL,
                          local_pis = '',
                          export = TRUE){
    
    x <- Map(format_xlsx, x, as.list(local_pis))
    tmp <- Map(xlsx_exporter, x, as.list(names(x)),
               as.list(export))
    
}



#' Create pdf randomization lists from a list of blockrand
#' generated data.frame 
#'
#' Create pdf randomization lists from a list of blockrand
#' generated data.frame
#' 
#' @param x a single data.frame or a named randlist (that is a
#'     data.frame with id and treatment columns). Names are used for
#'     file naming
#' @param path_prefix path prefix of the files to save in (overwriting
#'     the contents).
#' @param footer a character vector used as page central
#'     footer(s). Must be of length 1 if x is a data.frame or of the
#'     same length of x, if it's a list.
#' @param export logical recycled, wheter to export or not a list
#'     (used to export selectively)
#' @export
randlist2pdf <- function(x = NULL,
                         path_prefix = '/tmp/randlist',
                         footer = "",
                         export = TRUE) {

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
    Map(function(db, footer, file, export){
        if (export) make_randlist_pdf(x = db, cfoot = footer, f = file)
        else NULL
    }, x, as.list(footer), as.list(files), as.list(export))
    invisible(NULL)
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
#' @param export logical recycled, wheter to export or not a list
#'     (used to export selectively)
#' @param env_params list of parameters passed to blockrand::plotblockrand
#' @export
randlist2envelopes <- function(x = NULL,
                               path_prefix = '/tmp/envelopes',
                               study_acronym = "",
                               export = TRUE, 
                               env_params = list()) {

    ## make a list of data.frames
    x <- normalize_randlists(x)
    ## keep only exportable lists
    x <- x[export]
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
