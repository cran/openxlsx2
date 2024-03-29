% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/wb_styles.R
\name{create_border}
\alias{create_border}
\title{Helper to create a border}
\usage{
create_border(
  diagonal_down = "",
  diagonal_up = "",
  outline = "",
  bottom = NULL,
  bottom_color = NULL,
  diagonal = NULL,
  diagonal_color = NULL,
  end = "",
  horizontal = "",
  left = NULL,
  left_color = NULL,
  right = NULL,
  right_color = NULL,
  start = "",
  top = NULL,
  top_color = NULL,
  vertical = "",
  ...
)
}
\arguments{
\item{diagonal_down}{x}

\item{diagonal_up}{x}

\item{outline}{x}

\item{bottom}{X}

\item{bottom_color, diagonal_color, left_color, right_color, top_color}{a color created with \code{\link[=wb_color]{wb_color()}}}

\item{diagonal}{X}

\item{end}{x,}

\item{horizontal}{x}

\item{left}{x}

\item{right}{x}

\item{start}{x}

\item{top}{x}

\item{vertical}{x}

\item{...}{x}
}
\description{
Border styles can any of the following: "thin", "thick", "slantDashDot", "none", "mediumDashed", "mediumDashDot", "medium", "hair", "double", "dotted", "dashed", "dashedDotDot", "dashDot"
Border colors can be created with \code{\link[=wb_color]{wb_color()}}
}
\seealso{
\code{\link[=wb_add_border]{wb_add_border()}}

Other style creating functions: 
\code{\link{create_cell_style}()},
\code{\link{create_colors_xml}()},
\code{\link{create_dxfs_style}()},
\code{\link{create_fill}()},
\code{\link{create_font}()},
\code{\link{create_numfmt}()},
\code{\link{create_tablestyle}()}
}
\concept{style creating functions}
