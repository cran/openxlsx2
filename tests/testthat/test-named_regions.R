
test_that("Maintaining Named Regions on Load", {
  ## create named regions
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_worksheet("Sheet 2")

  ## specify region
  wb$add_data(sheet = 1, x = iris, startCol = 1, startRow = 1)
  wb$add_named_region(
    sheet = 1,
    name = "iris",
    rows = seq_len(nrow(iris) + 1),
    cols = seq_len(ncol(iris))
  )

  ## using write_data 'name' argument
  wb$add_data(sheet = 1, x = iris, name = "iris2", startCol = 10)

  ## Named region size 1
  wb$add_data(sheet = 2, x = 99, name = "region1", startCol = 3, startRow = 3)

  ## save file for testing
  out_file <- temp_xlsx()
  wb_save(wb, out_file, overwrite = TRUE)

  expect_equal(object = get_named_regions(wb), expected = get_named_regions(out_file))

  df1 <- read_xlsx(wb, namedRegion = "iris")
  df2 <- read_xlsx(out_file, namedRegion = "iris")
  expect_equal(df1, df2)

  df1 <- read_xlsx(wb, namedRegion = "region1", skipEmptyCols = FALSE)
  expect_s3_class(df1, "data.frame")
  expect_equal(nrow(df1), 0)
  expect_equal(ncol(df1), 1)

  df1 <- read_xlsx(wb, namedRegion = "region1", colNames = FALSE)
  expect_s3_class(df1, "data.frame")
  expect_equal(nrow(df1), 1)
  expect_equal(ncol(df1), 1)

  # nonsense
  # df1 is a single value and this single value is now used as rowName
  expect_warning(df1 <- read_xlsx(wb, namedRegion = "region1", rowNames = TRUE))
  expect_s3_class(df1, "data.frame")
  expect_equal(nrow(df1), 0)
  expect_equal(ncol(df1), 0)
})

test_that("Correctly Loading Named Regions Created in Excel", {

  # Load an excel workbook (in the repo, it's located in the /inst folder;
  # when installed on the user's system, it is located in the installation folder
  # of the package)
  filename <- system.file("extdata", "namedRegions.xlsx", package = "openxlsx2")

  # Load this workbook. We will test read_xlsx by passing both the object wb and
  # the filename. Both should produce the same results.
  wb <- wb_load(filename)

  # NamedTable refers to Sheet1!$C$5:$D$8
  table_f <- read_xlsx(filename,
    namedRegion = "NamedTable"
  )
  table_w <- read_xlsx(wb,
    namedRegion = "NamedTable"
  )

  expect_equal(object = table_f, expected = table_w)
  expect_equal(object = class(table_f), expected = "data.frame")
  expect_equal(object = ncol(table_f), expected = 2)
  expect_equal(object = nrow(table_f), expected = 3)

  # NamedCell refers to Sheet1!$C$2
  # This proeduced an error in an earlier version of the pacage when the object
  # wb was passed, but worked correctly when the filename was passed to read_xlsx
  cell_f <- read_xlsx(filename,
    namedRegion = "NamedCell",
    colNames = FALSE,
    rowNames = FALSE
  )

  cell_w <- read_xlsx(wb,
    namedRegion = "NamedCell",
    colNames = FALSE,
    rowNames = FALSE
  )

  expect_equal(object = cell_f, expected = cell_w)
  expect_equal(object = class(cell_f), expected = "data.frame")
  expect_equal(object = NCOL(cell_f), expected = 1)
  expect_equal(object = NROW(cell_f), expected = 1)

  # NamedCell2 refers to Sheet1!$C$2:$C$2
  cell2_f <- read_xlsx(filename,
    namedRegion = "NamedCell2",
    colNames = FALSE,
    rowNames = FALSE
  )

  cell2_w <- read_xlsx(wb,
    namedRegion = "NamedCell2",
    colNames = FALSE,
    rowNames = FALSE
  )

  expect_equal(object = cell2_f, expected = cell2_w)
  expect_equal(object = class(cell2_f), expected = "data.frame")
  expect_equal(object = NCOL(cell2_f), expected = 1)
  expect_equal(object = NROW(cell2_f), expected = 1)
})

# Ordering locally and in testthat differs.
test_that("Load names from an Excel file with funky non-region names", {
  filename <- system.file("extdata", "namedRegions2.xlsx", package = "openxlsx2")
  wb <- wb_load(filename)
  dn <- get_named_regions(wb)

  expect_equal(
    head(dn$name, 5),
    c("IQ_CH", "IQ_CQ", "IQ_CY", "IQ_DAILY", "IQ_FH")
  )
  expect_equal(
    dn$sheets,
    c(rep("", 26), "Sheet1",
      "Sheet with space", "Sheet1", "Sheet with space"
    )
  )
  expect_equal(dn$coords, c(rep("", 26), "B3", "B4", "B4", "B3"))

  dn2 <- get_named_regions(filename)
  expect_equal(dn, dn2)
})


test_that("Missing rows in named regions", {
  temp_file <- temp_xlsx()

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")

  ## create region
  wb$add_data(sheet = 1, x = iris[1:11, ], startCol = 1, startRow = 1)
  delete_data(wb, sheet = 1, cols = 1:2, rows = c(6, 6))

  wb$add_named_region(
    sheet = 1,
    name = "iris",
    rows = 1:(5 + 1),
    cols = 1:2
  )

  wb$add_named_region(
    sheet = 1,
    name = "iris2",
    rows = 1:(5 + 2),
    cols = 1:2
  )

  ## iris region is rows 1:6 & cols 1:2
  ## iris2 region is rows 1:7 & cols 1:2

  ## row 6 columns 1 & 2 are blank
  expect_equal(get_named_regions(wb)$name, c("iris", "iris2"))
  expect_equal(get_named_regions(wb)$sheets, c("Sheet 1", "Sheet 1"))
  expect_equal(get_named_regions(wb)$coords, c("A1:B6", "A1:B7"))

  ######################################################################## from Workbook

  ## Skip empty rows
  x <- read_xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(4, 2))

  x <- read_xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(5, 2))


  ## Keep empty rows
  x <- read_xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(5, 2))

  x <- read_xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(6, 2))



  ######################################################################## from file
  wb_save(wb, temp_file)

  ## Skip empty rows
  x <- read_xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(4, 2))

  x <- read_xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(5, 2))


  ## Keep empty rows
  x <- read_xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(5, 2))

  x <- read_xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(6, 2))

  unlink(temp_file)
})


test_that("Missing columns in named regions", {
  temp_file <- temp_xlsx()

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")

  ## create region
  wb$add_data(sheet = 1, x = iris[1:11, ], startCol = 1, startRow = 1)
  delete_data(wb, sheet = 1, cols = 2, rows = 1:12)

  wb$add_named_region(
    sheet = 1,
    name = "iris",
    rows = 1:5,
    cols = 1:2
  )

  wb$add_named_region(
    sheet = 1,
    name = "iris2",
    rows = 1:5,
    cols = 1:3
  )

  ## iris region is rows 1:5 & cols 1:2
  ## iris2 region is rows 1:5 & cols 1:3

  ## row 6 columns 1 & 2 are blank
  expect_equal(get_named_regions(wb)$name, c("iris", "iris2"), ignore_attr = TRUE)
  expect_equal(get_named_regions(wb)$sheets, c("Sheet 1", "Sheet 1"))
  expect_equal(get_named_regions(wb)$coords, c("A1:B5", "A1:C5"))

  ######################################################################## from Workbook

  ## Skip empty cols
  x <- read_xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 1))

  x <- read_xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 2))


  ## Keep empty cols
  x <- read_xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 2))

  x <- read_xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 3))



  ######################################################################## from file
  wb_save(wb, temp_file)

  ## Skip empty cols
  x <- read_xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 1))

  x <- read_xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 2))


  ## Keep empty cols
  x <- read_xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 2))

  x <- read_xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 3))

  unlink(temp_file)
})


test_that("Matching Substrings breaks reading named regions", {
  temp_file <- temp_xlsx()

  wb <- wb_workbook()
  wb$add_worksheet("table")
  wb$add_worksheet("table2")

  t1 <- head(iris)
  t1$Species <- as.character(t1$Species)
  t2 <- head(mtcars)

  wb$add_data(sheet = "table", x = t1, name = "t", startCol = 3, startRow = 12)
  wb$add_data(sheet = "table2", x = t2, name = "t2", startCol = 5, startRow = 24, rowNames = TRUE)

  wb$add_data(sheet = "table", x = head(t1, 3), name = "t1", startCol = 9, startRow = 3)
  wb$add_data(sheet = "table2", x = head(t2, 3), name = "t22", startCol = 15, startRow = 12, rowNames = TRUE)

  wb_save(wb, temp_file)

  r1 <- get_named_regions(wb)
  expect_equal(r1$sheets, c("table", "table", "table2", "table2"))
  expect_equal(r1$coords, c("C12:G18", "I3:M6", "E24:P30", "O12:Z15"))
  expect_equal(r1$name, c("t", "t1", "t2", "t22"))

  r2 <- get_named_regions(temp_file)
  expect_equal(r2$sheets, c("table", "table", "table2", "table2"))
  expect_equal(r1$coords, c("C12:G18", "I3:M6", "E24:P30", "O12:Z15"))
  expect_equal(r2$name, c("t", "t1", "t2", "t22"))


  ## read file named region
  expect_equal(t1, read_xlsx(xlsxFile = temp_file, namedRegion = "t"), ignore_attr = TRUE)
  expect_equal(t2, read_xlsx(xlsxFile = temp_file, namedRegion = "t2", rowNames = TRUE), ignore_attr = TRUE)
  expect_equal(head(t1, 3), read_xlsx(xlsxFile = temp_file, namedRegion = "t1"), ignore_attr = TRUE)
  expect_equal(head(t2, 3), read_xlsx(xlsxFile = temp_file, namedRegion = "t22", rowNames = TRUE), ignore_attr = TRUE)

  ## read Workbook named region
  expect_equal(t1, read_xlsx(xlsxFile = wb, namedRegion = "t"), ignore_attr = TRUE)
  expect_equal(t2, read_xlsx(xlsxFile = wb, namedRegion = "t2", rowNames = TRUE), ignore_attr = TRUE)
  expect_equal(head(t1, 3), read_xlsx(xlsxFile = wb, namedRegion = "t1"), ignore_attr = TRUE)
  expect_equal(head(t2, 3), read_xlsx(xlsxFile = wb, namedRegion = "t22", rowNames = TRUE), ignore_attr = TRUE)

  unlink(temp_file)
})


test_that("Read namedRegion from specific sheet", {

  filename <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")

  namedR <- "MyRange"
  sheets <- read_sheet_names(filename)

  # read the correct sheets
  expect_equal(data.frame(X1 = "S1A1", X2 = "S1B1", stringsAsFactors = FALSE), read_xlsx(filename, sheet = "Sheet1", namedRegion = namedR, rowNames = FALSE, colNames = FALSE), ignore_attr = TRUE)
  expect_equal(data.frame(X1 = "S2A1", X2 = "S2B1", stringsAsFactors = FALSE), read_xlsx(filename, sheet = which(sheets %in% "Sheet2"), namedRegion = namedR, rowNames = FALSE, colNames = FALSE), ignore_attr = TRUE)
  expect_equal(data.frame(X1 = "S3A1", X2 = "S3B1", stringsAsFactors = FALSE), read_xlsx(filename, sheet = "Sheet3", namedRegion = namedR, rowNames = FALSE, colNames = FALSE), ignore_attr = TRUE)

  # Warning: Workbook has no such named region. (Wrong namedRegion selected.)
  expect_error(read_xlsx(filename, sheet = "Sheet2", namedRegion = "MyRage", rowNames = FALSE, colNames = FALSE))

  # Warning: Workbook has no such named region on this sheet. (Correct namedRegion, but wrong sheet selected.)
  expect_error(read_xlsx(filename, sheet = "Sheet4", namedRegion = namedR, rowNames = FALSE, colNames = FALSE))
})


test_that("Overwrite and delete named regions", {
  temp_file <- temp_xlsx()

  wb <- wb_workbook()
  expect_null(get_named_regions(wb))

  wb$add_worksheet("Sheet 1")
  expect_null(get_named_regions(wb))

  ## create region
  wb$add_data(1, iris[1:11, ], startCol = 1, startRow = 1, name = "iris")
  exp <- data.frame(
    name   = "iris",
    value  = "'Sheet 1'!A1:E12",
    sheets = "Sheet 1",
    coords = "A1:E12",
    id     = 1L,
    local  = 0,
    sheet  = 1L
  )
  expect_identical(get_named_regions(wb), exp)

  # no overwrite
  expect_error(wb$add_data(1, iris[1:11, ], startCol = 1, startRow = 1, name = "iris"))

  expect_error(wb$add_named_region(1, name = "iris", rows = 1:5, cols = 1:2))

  # overwrite
  wb$add_named_region(1, name = "iris", rows = 1:5, cols = 1:2, overwrite = TRUE)

  exp <- data.frame(
    name   = "iris",
    # oh, why are these `'` and not `"`?
    value  = "'Sheet 1'!$A$1:$B$5",
    # and this doesn't have the `'`?
    sheets = "Sheet 1",
    coords = "A1:B5",
    id     = 1L,
    local  = 0,
    sheet  = 1L
  )

  # check modification
  expect_identical(get_named_regions(wb), exp)

  # delete name region
  wb$remove_named_region(name = "iris")
  expect_false("iris" %in% get_named_regions(wb)$name)

  wb$add_named_region(1, name = "iris", rows = 1:5, cols = 1:2)
  expect_identical(get_named_regions(wb), exp)

  # removing a worksheet removes the named region as well
  wb <- wb_remove_worksheet(wb, 1)
  expect_null(get_named_regions(wb))
})


test_that("deprecation test", {

    filename <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")
    expect_warning(wb_to_df(filename, definedName = "MyRange"),
                   "wb_to_df(definedName = .) is deprecated.  Use wb_to_df(named_region = .) instead",
                   fixed = TRUE)

})

test_that("load table", {

  wb <- wb_workbook()
  # add a table
  wb$add_worksheet("Sheet 1")
  wb$add_data_table(sheet = "Sheet 1", x = iris, tableName = "iris_tab")
  # add a named region
  wb$add_worksheet("Sheet 2")
  wb$add_data(sheet = "Sheet 2", x = iris, startCol = 1, startRow = 1)
  wb$add_named_region(
    sheet = 2,
    name = "iris",
    rows = seq_len(nrow(iris) + 1),
    cols = seq_along(iris)
  )

  expect_equal(c("iris", "iris_tab"), get_named_regions(wb)$name)

  expect_equal(
    wb_to_df(wb, named_region = "iris_tab"),
    wb_to_df(wb, named_region = "iris")
  )

})
