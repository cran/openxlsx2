% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/wb_load.R
\name{wb_load}
\alias{wb_load}
\title{Load an existing .xlsx, .xlsm or .xlsb file}
\usage{
wb_load(file, sheet, data_only = FALSE, ...)
}
\arguments{
\item{file}{A path to an existing .xlsx, .xlsm or .xlsb file}

\item{sheet}{optional sheet parameter. if this is applied, only the selected
sheet will be loaded. This can be a numeric, a string or \code{NULL}.}

\item{data_only}{mode to import if only a data frame should be returned. This
strips the \code{wbWorkbook} to a bare minimum.}

\item{...}{additional arguments}
}
\value{
A Workbook object.
}
\description{
\code{wb_load()} returns a \link{wbWorkbook} object conserving the content of the
original input file, including data, styles, media. This workbook can be
modified, read from, and be written back into a xlsx file.
}
\details{
If a specific \code{sheet} is selected, the workbook will still contain sheets
for all worksheets. The argument \code{sheet} and \code{data_only} are used internally
by \code{\link[=wb_to_df]{wb_to_df()}} to read from a file with minimal changes. They are not
specifically designed to create rudimentary but otherwise fully functional
workbooks. It is possible to import with
\code{wb_load(data_only = TRUE, sheet = NULL)}. In this way, only a workbook
framework is loaded without worksheets or data. This can be useful if only
some workbook properties are of interest.

There are some internal arguments that can be passed to wb_load, which are
used for development. The \code{debug} argument allows debugging of \code{xlsb} files
in particular. With \code{calc_chain} it is possible to maintain the calculation
chain. The calculation chain is used by spreadsheet software to determine
the order in which formulas are evaluated. Removing the calculation chain
has no known effect. The calculation chain is created the next time the
worksheet is loaded into the spreadsheet. Keeping the calculation chain
could only shorten the loading time in said software. Unfortunately, if a
cell is added to the worksheet, the calculation chain may block the
worksheet as the formulas will not be evaluated again until each individual
cell with a formula is selected in the spreadsheet software and the Enter
key is pressed manually. It is therefore strongly recommended not to
activate this function.

In rare cases, a warning is issued when loading an xlsx file that an xml
namespace has been removed from xml files. This refers to the internal
structure of the loaded xlsx file. Certain xlsx files created by third-party
applications contain a namespace (usually x). This namespace is not required
for the file to work in spreadsheet software and is not expected by
\code{openxlsx2}. It is therefore removed when the file is loaded into a
workbook. Removal is generally considered safe, but the feature is still not
commonly observed, hence the warning.

Initial support for binary openxml files (\code{xlsb}) has been added to the
package. We parse the binary file format into pseudo-openxml files that we
can import. Therefore, once imported, it is possible to interact with the
file as if it had been provided in xlsx file format in the first place. This
parsing into pseudo xml files is of course slower than reading directly from
the binary file. Our implementation is also still missing some functions:
some array formulas are not yet correct, conditional formatting and data
validation are not implemented, nor are pivot tables and slicers.
}
\examples{
## load existing workbook
fl <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
wb <- wb_load(file = fl)
}
