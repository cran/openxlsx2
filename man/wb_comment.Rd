% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/class-comment.R
\name{wb_comment}
\alias{wb_comment}
\title{Helper to create a comment object}
\usage{
wb_comment(
  text = NULL,
  style = NULL,
  visible = FALSE,
  author = getOption("openxlsx2.creator"),
  width = 2,
  height = 4
)
}
\arguments{
\item{text}{Comment text. Character vector. or a \code{\link[=fmt_txt]{fmt_txt()}} string.}

\item{style}{A Style object or list of style objects the same length as comment vector.}

\item{visible}{Is comment visible? Default: \code{FALSE}.}

\item{author}{Author of comment. A string. By default, will look at \code{options("openxlsx2.creator")}.
Otherwise, will check the system username.}

\item{width}{Textbox integer width in number of cells}

\item{height}{Textbox integer height in number of cells}
}
\value{
A \code{wbComment} object
}
\description{
Creates a \code{wbComment} object. Use with \code{\link[=wb_add_comment]{wb_add_comment()}} to add to a worksheet location.
}
\examples{
wb <- wb_workbook()
wb$add_worksheet("Sheet 1")

# write comment without author
c1 <- wb_comment(text = "this is a comment", author = "", visible = TRUE)
wb$add_comment(dims = "B10", comment = c1)

# Write another comment with author information
c2 <- wb_comment(text = "this is another comment", author = "Marco Polo")
wb$add_comment(sheet = 1, dims = "C10", comment = c2)

# write a styled comment with system author
s1 <- create_font(b = "true", color = wb_color(hex = "FFFF0000"), sz = "12")
s2 <- create_font(color = wb_color(hex = "FF000000"), sz = "9")
c3 <- wb_comment(text = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))

wb$add_comment(sheet = 1, dims = wb_dims(3, 6), comment = c3)
}
