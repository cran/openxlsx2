% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/class-hyperlink.R
\name{wbHyperlink}
\alias{wbHyperlink}
\title{R6 class for a Workbook Hyperlink}
\value{
A character vector of html if \code{is_external} is \code{TRUE}, otherwise \code{NULL}
}
\description{
R6 class for a Workbook Hyperlink

R6 class for a Workbook Hyperlink
}
\details{
A hyperlink
}
\section{Public fields}{
\if{html}{\out{<div class="r6-fields">}}
\describe{
\item{\code{ref}}{ref}

\item{\code{target}}{target}

\item{\code{location}}{location}

\item{\code{display}}{display}

\item{\code{is_external}}{is_external}
}
\if{html}{\out{</div>}}
}
\section{Methods}{
\subsection{Public methods}{
\itemize{
\item \href{#method-wbHyperlink-new}{\code{wbHyperlink$new()}}
\item \href{#method-wbHyperlink-to_xml}{\code{wbHyperlink$to_xml()}}
\item \href{#method-wbHyperlink-to_target_xml}{\code{wbHyperlink$to_target_xml()}}
\item \href{#method-wbHyperlink-clone}{\code{wbHyperlink$clone()}}
}
}
\if{html}{\out{<hr>}}
\if{html}{\out{<a id="method-wbHyperlink-new"></a>}}
\if{latex}{\out{\hypertarget{method-wbHyperlink-new}{}}}
\subsection{Method \code{new()}}{
Creates a new \code{wbHyperlink} object
\subsection{Usage}{
\if{html}{\out{<div class="r">}}\preformatted{wbHyperlink$new(ref, target, location, display = NULL, is_external = TRUE)}\if{html}{\out{</div>}}
}

\subsection{Arguments}{
\if{html}{\out{<div class="arguments">}}
\describe{
\item{\code{ref}}{ref}

\item{\code{target}}{target}

\item{\code{location}}{location}

\item{\code{display}}{display}

\item{\code{is_external}}{is_external}
}
\if{html}{\out{</div>}}
}
\subsection{Returns}{
a \code{wbHyperlink} object
}
}
\if{html}{\out{<hr>}}
\if{html}{\out{<a id="method-wbHyperlink-to_xml"></a>}}
\if{latex}{\out{\hypertarget{method-wbHyperlink-to_xml}{}}}
\subsection{Method \code{to_xml()}}{
Convert to xml
\subsection{Usage}{
\if{html}{\out{<div class="r">}}\preformatted{wbHyperlink$to_xml(id)}\if{html}{\out{</div>}}
}

\subsection{Arguments}{
\if{html}{\out{<div class="arguments">}}
\describe{
\item{\code{id}}{???}
}
\if{html}{\out{</div>}}
}
\subsection{Returns}{
A character vector of xml
}
}
\if{html}{\out{<hr>}}
\if{html}{\out{<a id="method-wbHyperlink-to_target_xml"></a>}}
\if{latex}{\out{\hypertarget{method-wbHyperlink-to_target_xml}{}}}
\subsection{Method \code{to_target_xml()}}{
Convert to target xml
\subsection{Usage}{
\if{html}{\out{<div class="r">}}\preformatted{wbHyperlink$to_target_xml(id)}\if{html}{\out{</div>}}
}

\subsection{Arguments}{
\if{html}{\out{<div class="arguments">}}
\describe{
\item{\code{id}}{???}
}
\if{html}{\out{</div>}}
}
}
\if{html}{\out{<hr>}}
\if{html}{\out{<a id="method-wbHyperlink-clone"></a>}}
\if{latex}{\out{\hypertarget{method-wbHyperlink-clone}{}}}
\subsection{Method \code{clone()}}{
The objects of this class are cloneable with this method.
\subsection{Usage}{
\if{html}{\out{<div class="r">}}\preformatted{wbHyperlink$clone(deep = FALSE)}\if{html}{\out{</div>}}
}

\subsection{Arguments}{
\if{html}{\out{<div class="arguments">}}
\describe{
\item{\code{deep}}{Whether to make a deep clone.}
}
\if{html}{\out{</div>}}
}
}
}
