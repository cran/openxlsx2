
# class -------------------------------------------------------------------

#' R6 class for a Workbook Worksheet
#'
#' A Worksheet
#'
#' @export
wbWorksheet <- R6::R6Class(
  "wbWorksheet",

  ## public ----
  public = list(

    # TODO can any of these be private?

    #' @field sheetPr sheetPr
    sheetPr = character(),

    #' @field dimension dimension
    dimension = character(),

    #' @field sheetViews sheetViews
    sheetViews = character(),

    #' @field sheetFormatPr sheetFormatPr
    sheetFormatPr = character(),

    #' @field sheet_data sheet_data
    sheet_data = NULL,

    #' @field cols_attr cols_attr
    cols_attr  = NULL,

    #' @field autoFilter autoFilter
    autoFilter = character(),

    #' @field mergeCells mergeCells
    mergeCells = NULL,

    #' @field conditionalFormatting conditionalFormatting
    conditionalFormatting = character(),

    #' @field dataValidations dataValidations
    dataValidations = NULL,

    #' @field freezePane freezePane
    freezePane = character(),

    #' @field hyperlinks hyperlinks
    hyperlinks = NULL,

    #' @field sheetProtection sheetProtection
    sheetProtection = character(),

    #' @field pageMargins pageMargins
    pageMargins = character(),

    #' @field pageSetup pageSetup
    pageSetup = character(),

    #' @field headerFooter headerFooter
    headerFooter = NULL,

    #' @field rowBreaks rowBreaks
    rowBreaks = character(),

    #' @field colBreaks colBreaks
    colBreaks = character(),

    #' @field drawing drawing
    drawing = character(),

    #' @field legacyDrawing legacyDrawing
    legacyDrawing = character(),

    #' @field legacyDrawingHF legacyDrawingHF
    legacyDrawingHF = character(),

    #' @field oleObjects oleObjects
    oleObjects = character(),

    #' @field tableParts tableParts
    tableParts = character(),

    #' @field extLst extLst
    extLst = character(),

    ### list with imported openxml-2.8.1 nodes
    #' @field cellWatches cellWatches
    cellWatches = character(),

    #' @field controls controls
    controls = character(),

    #' @field customProperties customProperties
    customProperties = character(),

    #' @field customSheetViews customSheetViews
    customSheetViews = character(),

    #' @field dataConsolidate dataConsolidate
    dataConsolidate = character(),

    #' @field drawingHF drawingHF
    drawingHF = character(),

    #' @field ignoredErrors ignoredErrors
    ignoredErrors = character(),

    #' @field phoneticPr phoneticPr
    phoneticPr = character(),

    #' @field picture picture
    picture = character(),

    #' @field printOptions printOptions
    printOptions = character(),

    #' @field protectedRanges protectedRanges
    protectedRanges = character(),

    #' @field scenarios scenarios
    scenarios = character(),

    #' @field sheetCalcPr sheetCalcPr
    sheetCalcPr = character(),

    #' @field smartTags smartTags
    smartTags = character(),

    #' @field sortState sortState
    sortState = character(),

    #' @field webPublishItems webPublishItems
    webPublishItems = character(),

    #' @description
    #' Creates a new `wbWorksheet` object
    #' @param tabColour tabColour
    #' @param oddHeader oddHeader
    #' @param oddFooter oddFooter
    #' @param evenHeader evenHeader
    #' @param evenFooter evenFooter
    #' @param firstHeader firstHeader
    #' @param firstFooter firstFooter
    #' @param paperSize paperSize
    #' @param orientation orientation
    #' @param hdpi hdpi
    #' @param vdpi vdpi
    #' @param printGridLines printGridLines
    #' @return a `wbWorksheet` object
    initialize = function(
      tabColour   = NULL,
      oddHeader   = NULL,
      oddFooter   = NULL,
      evenHeader  = NULL,
      evenFooter  = NULL,
      firstHeader = NULL,
      firstFooter = NULL,
      paperSize   = 9,
      orientation = "portrait",
      hdpi        = 300,
      vdpi        = 300,
      printGridLines = FALSE
    ) {
      if (!is.null(tabColour)) {
        tabColour <- sprintf('<sheetPr><tabColor rgb="%s"/></sheetPr>', tabColour)
      } else {
        tabColour <- character()
      }

      hf <- list(
        oddHeader   = na_to_null(oddHeader),
        oddFooter   = na_to_null(oddFooter),
        evenHeader  = na_to_null(evenHeader),
        evenFooter  = na_to_null(evenFooter),
        firstHeader = na_to_null(firstHeader),
        firstFooter = na_to_null(firstFooter)
      )

      if (all(lengths(hf) == 0)) {
        hf <- list()
      }

      # only add if printGridLines not TRUE. The openxml default is TRUE
      if (printGridLines) {
       self$set_print_options(gridLines = printGridLines, gridLinesSet = printGridLines)
      }

      ## list of all possible children
      self$sheetPr               <- tabColour
      self$dimension             <- '<dimension ref="A1"/>'
      self$sheetViews            <- character()
      self$sheetFormatPr         <- '<sheetFormatPr baseColWidth="8.43" defaultRowHeight="16" x14ac:dyDescent="0.2"/>'
      self$cols_attr             <- character()
      self$autoFilter            <- character()
      self$mergeCells            <- character()
      self$conditionalFormatting <- character()
      self$dataValidations       <- NULL
      self$hyperlinks            <- list()
      self$pageMargins           <- '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
      self$pageSetup             <- sprintf('<pageSetup paperSize="%s" orientation="%s" horizontalDpi="%s" verticalDpi="%s"/>', paperSize, orientation, hdpi, vdpi)
      self$headerFooter          <- hf
      self$rowBreaks             <- character()
      self$colBreaks             <- character()
      self$drawing               <- character()
      self$legacyDrawing         <- character()
      self$legacyDrawingHF       <- character()
      self$oleObjects            <- character()
      self$tableParts            <- character()
      self$extLst                <- character()
      self$freezePane            <- character()
      self$sheet_data            <- wbSheetData$new()

      invisible(self)
    },

    #' @description
    #' Get prior sheet data
    #' @return A character vector of xml
    get_prior_sheet_data = function() {

      # apparently every sheet needs to have a sheetView
      sheetViews <- self$sheetViews

      if (length(self$freezePane)) {
        if (length(xml_node(sheetViews, "sheetViews", "sheetView")) == 1) {
          # get sheetView node and append freezePane
          # TODO Can we unfreeze a pane? It should be possible to simply null freezePane
          sheetViews <- xml_add_child(sheetViews, xml_child = self$freezePane, level = "sheetView")
        } else {
          message("Sheet contains multiple sheetViews. Could not freeze pane") #nocov
        }
      }

      paste_c(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">',

        # sheetPr
        if (length(self$sheetPr) && !any(xml_node_name(self$sheetPr) == "sheetPr")) {
          xml_node_create("sheetPr", xml_children = self$sheetPr)
        } else {
          self$sheetPr
        },

        self$dimension,

        sheetViews,

        self$sheetFormatPr,
        # cols_attr
        # is this fine if it's just <cols></cols>?
        if (length(self$cols_attr)) {
          paste(c("<cols>", self$cols_attr, "</cols>"), collapse = "")
        },
        '</worksheet>',
        sep = ""
      )
    },

    #' @description
    #' Get post sheet data
    #' @return A character vector of xml
    get_post_sheet_data = function() {
      paste_c(
        self$sheetProtection,
        self$autoFilter,

        # mergeCells
        if (length(self$mergeCells)) {
          paste0(
            sprintf('<mergeCells count="%i">', length(self$mergeCells)),
            pxml(self$mergeCells),
            "</mergeCells>"
          )
        },

        # conditionalFormatting
        if (length(self$conditionalFormatting)) {
          nms <- names(self$conditionalFormatting)
          paste(
            vapply(
              unique(nms),
              function(i) {
                paste0(
                  sprintf('<conditionalFormatting sqref="%s">', i),
                  pxml(self$conditionalFormatting[nms == i]),
                  "</conditionalFormatting>"
                )
              },
              NA_character_
            ),
            collapse = ""
          )
        },

        # dataValidations
        if (length(self$dataValidations)) {
          paste0(
            sprintf('<dataValidations count="%i">', length(self$dataValidations)),
            pxml(self$dataValidations),
            "</dataValidations>"
          )
        },

        # hyperlinks
        if (n <- length(self$hyperlinks)) {
          h_inds <- paste0(seq_len(n), "h")
          paste(
            "<hyperlinks>",
            paste(
              vapply(
                seq_along(h_inds),
                function(i)  {
                  self$hyperlinks[[i]]$to_xml(h_inds[i])
                },
                NA_character_
              ),
              collapse = ""
            ),
            "</hyperlinks>"
          )
        },

        self$printOptions,
        self$pageMargins,
        self$pageSetup,

        # headerFooter
        # should return NULL when !length(self$headerFooter)
        genHeaderFooterNode(self$headerFooter),

        # rowBreaks
        if (n <- length(self$rowBreaks)) {
          paste0(
            sprintf('<rowBreaks count="%i" manualBreakCount="%i">', n, n),
            paste(self$rowBreaks, collapse = ""),
            "</rowBreaks>"
          )
        },

        # colBreaks
        if (n <- length(self$colBreaks)) {
          paste0(
            sprintf('<colBreaks count="%i" manualBreakCount="%i">', n, n),
            paste(self$colBreaks, collapse = ""),
            "</colBreaks>"
          )
        },

        self$drawing,
        self$legacyDrawing,
        self$legacyDrawingHF,
        self$oleObjects,

        # tableParts
        if (n <- length(self$tableParts)) {
          paste0(sprintf('<tableParts count="%i">', n), pxml(self$tableParts), "</tableParts>")
        },

        # extLst
        if (length(self$extLst)) {
          sprintf(
            "<extLst>%s</extLst>",
            paste0(
              pxml(self$extLst)
            )
          )
        },

        # end
        sep = ""
      )

    },

    #' @description
    #' unfold `<cols ..>` node to dataframe. `<cols><col ..>` are compressed.
    #' Only columns with attributes are written to the file. This function
    #' unfolds them so that each cell beginning with the "A" to the last one
    #' found in cc gets a value.
    #' TODO might extend this to match either largest cc or largest col. Could
    #' be that "Z" is formatted, but the last value is written to "Y".
    #' TODO might replace the xml nodes with the data frame?
    #' @return The column data frame
    unfold_cols = function() {

      # avoid error and return empty data frame
      if (length(self$cols_attr) == 0)
        return(empty_cols_attr())

      col_df <- col_to_df(read_xml(self$cols_attr))
      col_df$min <- as.numeric(col_df$min)
      col_df$max <- as.numeric(col_df$max)

      max_col <- max(col_df$max)

      # always begin at 1, even if 1 is not in the dataset. fold_cols requires this
      key <- seq(1, max_col)

      # merge against this data frame
      tmp_col_df <- data.frame(
        key = key,
        stringsAsFactors = FALSE
      )

      out <- NULL
      for (i in seq_len(nrow(col_df))) {
        z <- col_df[i, ]
        for (j in seq(z$min, z$max)) {
          z$key <- j
          out <- rbind(out, z)
        }
      }

      # merge and convert to character, remove key
      col_df <- merge(x = tmp_col_df, y = out, by = "key", all.x = TRUE)
      col_df$min <- as.character(col_df$key)
      col_df$max <- as.character(col_df$key)
      col_df[is.na(col_df)] <- ""
      col_df$key <- NULL

      col_df
    },

    #' @description
    #' fold the column dataframe back into a node.
    #' @param col_df the column data frame
    #' @return The `wbWorksheetObject`, invisibly
    fold_cols = function(col_df) {

      # remove min and max columns and create merge identifier: string
      col_df <- col_df[-which(names(col_df) %in% c("min", "max"))]
      col_df$string <- apply(col_df, 1, paste, collapse = "")

      # run length
      out <- with(
        rle(col_df$string),
        data.frame(
          string = values,
          min = cumsum(lengths) - lengths + 1,
          max = cumsum(lengths))
      )

      # remove duplicates pre merge
      col_df <- unique(col_df)

      # merge with string variable, drop empty string and clean up
      col_df <- merge(out, col_df, by = "string", all.x = TRUE)
      col_df <- col_df[col_df$string != "", ]
      col_df$string <- NULL

      # order and return
      col_df <- col_df[order(col_df$min), ]
      col_df$min <- as.character(col_df$min)
      col_df$max <- as.character(col_df$max)

      # assign as xml-nodes
      self$cols_attr <- df_to_xml("col", col_df)

      invisible(self)
    },


    #' @description clean sheet (remove all values)
    #' @param numbers remove all numbers
    #' @param characters remove all characters
    #' @param styles remove all styles
    #' @param merged_cells remove all merged_cells
    #' @return The `wbWorksheetObject`, invisibly
    clean_sheet = function(numbers = TRUE, characters = TRUE, styles = TRUE, merged_cells = TRUE) {

      cc <- self$sheet_data$cc

      if (NROW(cc) == 0) return(invisible(self))

      if (numbers)
        cc[cc$c_t %in% c("n", ""),
          c("c_t", "v", "f", "f_t", "f_ref", "f_ca", "f_si", "is")] <- ""

      if (characters)
        cc[cc$c_t %in% c("inlineStr", "s"),
          c("v", "f", "f_t", "f_ref", "f_ca", "f_si", "is")] <- ""

      if (styles)
        cc[c("c_s")] <- ""

      self$sheet_data$cc <- cc

      if (merged_cells)
        self$mergeCells <- character(0)

      invisible(self)

    },

    #' @description add page break
    #' @param row row
    #' @param col col
    #' @returns The `wbWorksheet` object
    add_page_break = function(row = NULL, col = NULL) {
      if (!xor(is.null(row), is.null(col))) {
        stop("either `row` or `col` must be NULL but not both")
      }

      if (!is.null(row)) {
        if (!is.numeric(row)) stop("`row` must be numeric")
        self$append("rowBreaks", sprintf('<brk id="%i" max="16383" man="1"/>', round(row)))
      } else if (!is.null(col)) {
        if (!is.numeric(col)) stop("`col` must be numeric")
        self$append("colBreaks", sprintf('<brk id="%i" max="1048575" man="1"/>', round(col)))
      }

      invisible(self)
    },

    #' @description add print options
    #' @param gridLines gridLines
    #' @param gridLinesSet gridLinesSet
    #' @param headings If TRUE prints row and column headings
    #' @param horizontalCentered If TRUE the page is horizontally centered
    #' @param verticalCentered If TRUE the page is vertically centered
    #' @returns The `wbWorksheet` object
    set_print_options = function(
        gridLines          = NULL,
        gridLinesSet       = NULL,
        headings           = NULL,
        horizontalCentered = NULL,
        verticalCentered   = NULL
    ) {
      self$printOptions <- xml_node_create(
        xml_name = "printOptions",
        xml_attributes = c(
          gridLines          = as_xml_attr(gridLines),
          gridLinesSet       = as_xml_attr(gridLinesSet),
          headings           = as_xml_attr(headings),
          horizontalCentered = as_xml_attr(horizontalCentered),
          verticalCentered   = as_xml_attr(verticalCentered)
        )
      )
    },

    #' @description append a field.  Intended for internal use only.  Not
    #'   guaranteed to remain a public method.
    #' @param field a field name
    #' @param value a new value
    #' @return The `wbWorksheetObject`, invisibly
    append = function(field, value = NULL) {
      self[[field]] <- c(self[[field]], value)
      invisible(self)
    },

    #' @description add sparkline
    #' @param sparklines sparkline created by `create_sparkline()`
    #' @return The `wbWorksheetObject`, invisibly
    add_sparklines = function(
      sparklines
    ) {

      private$do_append_x14(sparklines, "x14:sparklineGroup", "x14:sparklineGroups")

      invisible(self)
    },

    #' @description add sheetview
    #' @param colorId colorId
    #' @param defaultGridColor defaultGridColor
    #' @param rightToLeft rightToLeft
    #' @param showFormulas showFormulas
    #' @param showGridLines showGridLines
    #' @param showOutlineSymbols showOutlineSymbols
    #' @param showRowColHeaders showRowColHeaders
    #' @param showRuler showRuler
    #' @param showWhiteSpace showWhiteSpace
    #' @param showZeros showZeros
    #' @param tabSelected tabSelected
    #' @param topLeftCell topLeftCell
    #' @param view view
    #' @param windowProtection windowProtection
    #' @param workbookViewId workbookViewId
    #' @param zoomScale zoomScale
    #' @param zoomScaleNormal zoomScaleNormal
    #' @param zoomScalePageLayoutView zoomScalePageLayoutView
    #' @param zoomScaleSheetLayoutView zoomScaleSheetLayoutView
    #' @return The `wbWorksheetObject`, invisibly
    set_sheetview = function(
      colorId                  = NULL,
      defaultGridColor         = NULL,
      rightToLeft              = NULL,
      showFormulas             = NULL,
      showGridLines            = NULL,
      showOutlineSymbols       = NULL,
      showRowColHeaders        = NULL,
      showRuler                = NULL,
      showWhiteSpace           = NULL,
      showZeros                = NULL,
      tabSelected              = NULL,
      topLeftCell              = NULL,
      view                     = NULL,
      windowProtection         = NULL,
      workbookViewId           = NULL,
      zoomScale                = NULL,
      zoomScaleNormal          = NULL,
      zoomScalePageLayoutView  = NULL,
      zoomScaleSheetLayoutView = NULL
    ) {

      # all zoom scales must be in the range of 10 - 400

      # get existing sheetView
      sheetView <- xml_node(self$sheetViews, "sheetViews", "sheetView")

      if (length(sheetView) == 0)
        sheetView <- xml_node_create("sheetView")

      sheetView <- xml_attr_mod(
        sheetView,
        xml_attributes = c(
          colorId                  = as_xml_attr(colorId),
          defaultGridColor         = as_xml_attr(defaultGridColor),
          rightToLeft              = as_xml_attr(rightToLeft),
          showFormulas             = as_xml_attr(showFormulas),
          showGridLines            = as_xml_attr(showGridLines),
          showOutlineSymbols       = as_xml_attr(showOutlineSymbols),
          showRowColHeaders        = as_xml_attr(showRowColHeaders),
          showRuler                = as_xml_attr(showRuler),
          showWhiteSpace           = as_xml_attr(showWhiteSpace),
          showZeros                = as_xml_attr(showZeros),
          tabSelected              = as_xml_attr(tabSelected),
          topLeftCell              = as_xml_attr(topLeftCell),
          view                     = as_xml_attr(view),
          windowProtection         = as_xml_attr(windowProtection),
          workbookViewId           = as_xml_attr(workbookViewId),
          zoomScale                = as_xml_attr(zoomScale),
          zoomScaleNormal          = as_xml_attr(zoomScaleNormal),
          zoomScalePageLayoutView  = as_xml_attr(zoomScalePageLayoutView),
          zoomScaleSheetLayoutView = as_xml_attr(zoomScaleSheetLayoutView)
        )
      )

      self$sheetViews <- xml_node_create(
        "sheetViews",
        xml_children = sheetView
      )

      invisible(self)
    }
  ),

  ## private ----
  private = list(
    # These were commented out in the RC object -- not sure if they're needed
    cols                  = NULL,
    sheetData             = NULL,

    # @description add data_validation_lst
    # @param datavalidation datavalidation
    do_append_x14 = function(
      x,
      s_name,
      l_name
    ) {

      if (!all(xml_node_name(x) == s_name))
        stop(sprintf("all nodes must match %s. Got %s", s_name, xml_node_name(x)))

      # can have length > 1 for multiple xmlns attributes. we take this extLst,
      # inspect it, update if needed and return it
      extLst <- xml_node(self$extLst, "ext")
      is_xmlns_x14 <- grepl(pattern = "xmlns:x14", extLst)

      # different ext types have different uri ids. We support dataValidations
      # and sparklineGroups.
      uri <- ""
      # if (l_name == "x14:dataValidations") uri <- "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"
      if (l_name == "x14:sparklineGroups") uri <- "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"

      is_needed_uri <- grepl(pattern = uri, extLst, fixed = TRUE)

      # check if any <ext xmlns:x14 ...> node exists, else add it
      if (length(extLst) == 0 || !any(is_xmlns_x14) || !any(is_needed_uri)) {
        ext <- xml_node_create(
          "ext",
          xml_attributes = c("xmlns:x14" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                             uri = uri)
        )

        # update extLst
        extLst <- c(extLst, ext)
        is_needed_uri <- c(is_needed_uri, TRUE)
      } else {
        ext <- extLst[is_needed_uri]
      }

      # check again and should be exactly one ext node
      is_xmlns_x14 <- grepl(pattern = "xmlns:x14", extLst)

      # check for l_name and add one if none is found
      if (length(xml_node(ext, "ext", l_name)) == 0) {
        ext <- xml_add_child(
          ext,
          xml_node_create(
            l_name,
            xml_attributes = c("xmlns:xm" = "http://schemas.microsoft.com/office/excel/2006/main"))
        )
      }

      # add new x to exisisting l_name
      ext <- xml_add_child(
        ext,
        level = c(l_name),
        x
      )

      # update counts for dataValidations
      # count is all matching nodes. not sure if required
      if (l_name == "x14:dataValidations") {

        outer <- xml_attr(ext, "ext")
        inner <- getXMLPtr1con(read_xml(ext))

        xdv <- grepl(l_name, inner)
        inner <- xml_attr_mod(
          inner[xdv],
          xml_attributes = c(count = as.character(length(xml_node_name(inner[xdv], l_name))))
        )

        ext <- xml_node_create("ext", xml_children = inner, xml_attributes = unlist(outer))
      }

      # update extLst and add it back to worksheet
      extLst[is_needed_uri] <- ext
      self$extLst <- extLst

      invisible(self)
    },

    data_validation = function(
      type,
      operator,
      value,
      allowBlank,
      showInputMsg,
      showErrorMsg,
      errorStyle,
      errorTitle,
      error,
      promptTitle,
      prompt,
      origin,
      sqref
    ) {

      header <- xml_node_create(
        "dataValidation",
        xml_attributes = c(
          type = type,
          operator = operator,
          allowBlank = allowBlank,
          showInputMessage = showInputMsg,
          showErrorMessage = showErrorMsg,
          sqref = sqref,
          errorStyle = errorStyle,
          errorTitle = errorTitle,
          error = error,
          promptTitle = promptTitle,
          prompt = prompt
        )
      )

      if (type == "date") {
        value <- as.integer(value) + origin
      }

      if (type == "time") {
        t <- format(value[1], "%z")
        offSet <-
          suppressWarnings(
            ifelse(substr(t, 1, 1) == "+", 1L, -1L) * (
              as.integer(substr(t, 2, 3)) + as.integer(substr(t, 4, 5)) / 60
            ) / 24
          )
        if (is.na(offSet)) {
          offSet[i] <- 0
        }

        value <- as.numeric(as.POSIXct(value)) / 86400 + origin + offSet
      }

      form <- sapply(
        seq_along(value),
        function(i) {
          sprintf("<formula%s>%s</formula%s>", i, value[i], i)
        }
      )

      self$append("dataValidations", xml_add_child(header, form))
      invisible(self)
    }
  )
)



wb_worksheet <- function() {
  wbWorksheet$new()
}

empty_cols_attr <- function(n = 0, beg, end) {
  # make make this a specific class/object?

  if (!missing(beg) && !missing(end)) {
    n_seq <- seq.int(beg, end, by = 1)
    n <- length(n_seq)
  } else {
     n_seq <- seq_len(n)
  }

  cols_attr_nams <- c("bestFit", "collapsed", "customWidth", "hidden", "max",
                      "min", "outlineLevel", "phonetic", "style", "width")

  z <- data.frame(
    matrix("", nrow = n, ncol = length(cols_attr_nams)),
    stringsAsFactors = FALSE
  )
  names(z) <- cols_attr_nams

  if (n > 0) {
    z$min <- n_seq
    z$max <- n_seq
    z$width <- "8.43"
  }

  z
}
