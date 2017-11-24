#' Convert Excel cell index to numeric column and row indexes
#'
#' This function converts an Execl cell index (with max 2 column letters) to row and column numbers.
#' @param celIndex with maxmum 2-letter column and a number, e.g. Ad23.
#' @param rrw a (+/-) integer relative to celIndex row, defaults to 0.
#' @param rcn a (+/-) integer relative to celIndex column, defaults to 0.
#' @return The number of row with cel2num(...)$r or cel2num(...)[1], and the number of column with cel2num(...)$c or cel2num(...)[2].
#' @keywords Excel
#' @export
#' @examples
#' cel2num("mD263")$r or cel2num("mD263")[1] for row number
#' cel2num("mD263")$c or cel2num("mD263")[2] for column number

cel2num <- function(celIndex, rrw = 0, rcn = 0) {
  rw <- gsub("[^[:digit:]]", "", celIndex)
  rw <- as.numeric(rw)
  
  chars <- gsub("[[:digit:]]", "", celIndex)
  chars <- toupper(chars)  #change to UPPER cases
  
  if (nchar(substr(chars, 2, 2))) {
    res <- which(LETTERS == substr(chars, 1, 1)) * 26 + which(LETTERS == 
      substr(chars, 2, 2))
  } else {
    res <- which(LETTERS == substr(chars, 1, 1))
  }
  
  CeL <- list(r = rw + rrw, c = res + rcn)
  return(CeL)  #return(CeL) can be also turned off, -if it is on, both objects will be printed out when function called.
  # The objects can be called with: cel2num(...)$r or rdxls(...)$c
}  #close function