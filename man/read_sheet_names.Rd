% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/readWorkbook.R
\name{read_sheet_names}
\alias{read_sheet_names}
\title{Get names of worksheets}
\usage{
read_sheet_names(file)
}
\arguments{
\item{file}{An xlsx or xlsm file.}
}
\value{
Character vector of worksheet names.
}
\description{
Returns the worksheet names within an xlsx file
}
\examples{
read_sheet_names(system.file("extdata", "readTest.xlsx", package = "openxlsx2"))
}
