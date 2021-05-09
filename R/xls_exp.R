#' Exporting data into .xlsx files
#'
#' Exporting one or multiple tables into a .xlsx file.
#' If there are multiple tables, each one will be stored in the same file in
#' different sheets.
#'
#' @param data data frame or list of data frames
#' @param sheetnames character vector used to name each file sheet
#' @param filename character to name the exported file
#' @param path path where to store the file inside the working directory
#' @param keepNA logical, indicating if missing values should be displayed or
#' not. If TRUE, are displayed. Default set to FALSE. Note: if keepNA = TRUE,
#' and the column is numeric, it would turn into character.
#' @param colnames logical, indicating if column names should be in the output. If
#' true, column names are in the first row. Default set to TRUE.
#'
#' @return .xlsx file exported in the indicated path.
#' @export
#'
#' @examples
xlsx_exp =
  function(
    data,
    sheetnames = NULL,
    filename,
    path,
    keepNA = F,
    colnames = T,
    rownames = F
  ) {

    # data parameter checking:
    ## To list class: needed to use every list element as a
    ## workbook sheet
    if (inherits(data, "list")) {
      dt = data
    } else {
      dt = list(data)
    }

    # sheetnames parameter:
    ## There are three options:
    ### 1. list has its own names and sheetnames were given
    ### 2. list hasn't its own names and sheetnames weren't given
    ### 3. list hasn't its own names and sheetnames were given
    ### 4. list has its own names and sheetnames weren't given
    if (is.null(names(dt))) {
      if (is.null(sheetnames)) {
        # Option 2:
        sheetnames = paste0("Sheet", seq(1:length(dt)))
      } else {
        # Option 3
        sheetnames = sheetnames
      }
    } else {
      if (is.null(sheetnames)) {
        # Option 4
        sheetnames = names(dt)
      } else {
        # Option 1
        sheetnames = sheetnames
        warning("Although data has it's own names, user given sheetnames were used instead")
      }
    }

    # Keep NA's or not
    if (keepNA) {
      kNA = T
      message("Missing values will be displayed as NA")
    } else {
      kNA = F
      message("Missing values will be empty")
    }

    # Defaults about colnames and rownames
    if (colnames == T) {
      CN = T
    } else {
      CN = F
    }

    if (rownames == T) {
      RN = T
    } else {
      RN = F
    }

    # Creating the workbook
    Wb = openxlsx::createWorkbook(".xlsx")

    # Adding each sheet
    for (i in 1:length(sheetnames)) {
      openxlsx::addWorksheet(
        wb = Wb, sheetName = sheetnames[i]
      )
      openxlsx::writeData(
        wb = Wb,
        sheet =  i,
        x = dt[[i]],
        colNames = CN,
        rowNames = RN,
        keepNA = kNA
      )
    }

    # Saving the file
    if (missing(path)) {
      openxlsx::saveWorkbook(
        wb = Wb,
        file = paste0(filename,".xlsx" ),
        overwrite = T
      )
      message("Since there is not path, file saved in working directory")

    } else {
      openxlsx::saveWorkbook(
        wb = Wb,
        file = paste0("./",path,"/",filename,".xlsx" ),
        overwrite = T
      )
    }
  }
