% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/wb_styles.R
\name{create_font}
\alias{create_font}
\title{create font format}
\usage{
create_font(
  b = "",
  charset = "",
  color = wb_colour(hex = "FF000000"),
  condense = "",
  extend = "",
  family = "2",
  i = "",
  name = "Calibri",
  outline = "",
  scheme = "minor",
  shadow = "",
  strike = "",
  sz = "11",
  u = "",
  vertAlign = ""
)
}
\arguments{
\item{b}{bold}

\item{charset}{charset}

\item{color}{rgb color: default "FF000000"}

\item{condense}{condense}

\item{extend}{extend}

\item{family}{font family: default "2"}

\item{i}{italic}

\item{name}{font name: default "Calibri"}

\item{outline}{outline}

\item{scheme}{font scheme: default "minor"}

\item{shadow}{shadow}

\item{strike}{strike}

\item{sz}{font size: default "11",}

\item{u}{underline}

\item{vertAlign}{vertical alignment}
}
\description{
create font format
}
\examples{
font <- create_font()
# openxml has the alpha value leading
hex8 <- unlist(xml_attr(read_xml(font), "font", "color"))
hex8 <- paste0("#", substr(hex8, 3, 8), substr(hex8, 1,2))

# # write test color
# col <- crayon::make_style(col2rgb(hex8, alpha = TRUE))
# cat(col("Test"))
}
