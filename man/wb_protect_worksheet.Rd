% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/class-workbook-wrappers.R
\name{wb_protect_worksheet}
\alias{wb_protect_worksheet}
\title{Protect a worksheet from modifications}
\usage{
wb_protect_worksheet(
  wb,
  sheet = current_sheet(),
  protect = TRUE,
  password = NULL,
  properties = NULL
)
}
\arguments{
\item{wb}{A workbook object}

\item{sheet}{A name or index of a worksheet}

\item{protect}{Whether to protect or unprotect the sheet (default=TRUE)}

\item{password}{(optional) password required to unprotect the worksheet}

\item{properties}{A character vector of properties to lock.  Can be one or
more of the following: \code{"selectLockedCells"}, \code{"selectUnlockedCells"},
\code{"formatCells"}, \code{"formatColumns"}, \code{"formatRows"}, \code{"insertColumns"},
\code{"insertRows"}, \code{"insertHyperlinks"}, \code{"deleteColumns"}, \code{"deleteRows"},
\code{"sort"}, \code{"autoFilter"}, \code{"pivotTables"}, \code{"objects"}, \code{"scenarios"}}
}
\description{
Protect or unprotect a worksheet from modifications by the user in the graphical user interface. Replaces an existing protection. Certain features require applying unlocking of initialized cells in the worksheet and across columns and/or rows.
}
\examples{
wb <- wb_workbook()
wb$add_worksheet("S1")
wb$add_data_table(1, x = iris[1:30, ])

wb$protect_worksheet(
  "S1",
  protect = TRUE,
  properties = c("formatCells", "formatColumns", "insertColumns", "deleteColumns")
)

# Formatting cells / columns is allowed , but inserting / deleting columns is protected:
wb$protect_worksheet(
  "S1",
  protect = TRUE,
   c(formatCells = FALSE, formatColumns = FALSE,
                 insertColumns = TRUE, deleteColumns = TRUE)
)

# Remove the protection
wb$protect_worksheet("S1", protect = FALSE)
}
