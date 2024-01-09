#' IOPS
#'
#' @description  Takes user inputted trade data, any acceptable ISO country code and industrial value chain mapping to calculate various metrics (Economic- and Product complexity metrics, distance metrics, opportunity gain, and inequality metrics) of a given country in order to facilitate better decision making regarding industrial policymaking.
#'
#' @import dplyr
#' @import tidyr
#' @import readxl
#' @import economiccomplexity
#' @import usethis
#' @importFrom stats aggregate
#' @importFrom utils write.csv
#' @importFrom utils read.csv
#' @importFrom openxlsx createWorkbook 
#' @importFrom openxlsx addWorksheet 
#' @importFrom openxlsx writeData
#' @importFrom openxlsx saveWorkbook 
#'
#'
#' @param CountryCode (Type: character/integer) Any accepted ISO country code could be used, e.g. \code{"United Kingdom"}, \code{"GBR"}, \code{"GB"}, \code{828} would all be accepted if the United Kingdom is the desired country. This version of the package uses country_codes_V202201 from th BACI Hs17 2020 dataset, but can be changed in \code{inst/extdata}.
#' @param tradeData (Type: csv) Accepts any CEPII BACI trade data. NOTE: tradeData and GVCMapping must be from the same "H" Family, e.g. both are from  H3, etc., in order for the program to work correctly.
#' @param ComplexMethod (Type: character) Methods used to calculate complexity measures. Can be any one of these methods: \code{"fitness"}, \code{"reflections"} or \code{"eigenvalues"}. Defaults to "eigenvalues".
#' @param iterCompl (Type: integer) The number of iterations that the chosen complexity measure must use. Defaults to \code{iterCompl = 20}.
#' @param GVCMapping (Type: csv) The desired value chain to be analysed. With Columns "Tiers", "Activity", and "HSCode". NOTE: tradeData and GVCMapping must be from the same "H" Family, e.g. both are from  H3, etc., in order for the program to work correctly.
#' @param tradedigit (Type: integer) Indicate if the raw trade digit summation should be done on a 6- or 4-digit level. Defaults to tradedigit = 6.
#'
#' @examples
#' 
#' # Create a temporary directory
#' temp_dir <- tempfile()
#' dir.create(temp_dir)
#' 
#' # Set the working directory to the temporary directory
#' old_dir <- setwd(temp_dir)
#' 
#' #Generate example trade data
#' GeneratedTradeData <- data.frame(
#'   t = c(2020, 2020, 2020),
#'   i = c(4, 710, 710),
#'   j = c(842, 124, 251),
#'   k = c(842410, 110220, 845210),
#'   v = c(4.776, 0.088, 0.057),
#'   q = c(0.025, 0.007, 0.005)
#' )
#'
#' # Use your function with generated trade data
#' IOPS(
#'   CountryCode = 710,
#'   tradeData = GeneratedTradeData,
#'   ComplexMethod = "reflections",
#'   iterCompl = 2,
#'   GVCMapping = NULL,
#'   tradedigit = 6
#' )
#' 
#' \donttest{
#' # Use your function with real trade data
#' IOPS(
#'   CountryCode = 710,
#'   tradeData = ExampleTradeData,
#'   ComplexMethod = "reflections",
#'   iterCompl = 2,
#'   GVCMapping = NULL,
#'   tradedigit = 6
#' )
#' }
#'
#' # Clean up the temporary directory
#'   setwd(old_dir)  # Restore the original working directory
#'   unlink(temp_dir, recursive = TRUE)
#' 
#'
#' @return A list that constrains ECI, PCI, Opportunity_Gain, distance, density, M_absolute, M_binary, Tier_Results, Product_Category_Results, Product_Results, respectively.
#' @export

IOPS <- function(CountryCode ,tradeData , ComplexMethod = "eigenvalues", iterCompl = 20, GVCMapping = NULL, tradedigit = 6){

  #Assign empty variables
  country_code <- NA
  country_name_abbreviation <- NA
  country_name_full <- NA
  iso_2digit_alpha <- NA
  iso_3digit_alpha <- NA
  hs_product_code <- NA
  export_value <- NA
  GVCactivityNumber <- NA
  tierNumber <- NA
  i <- NA
  k <- NA
  v <- NA
  HScode <- NA
  location_code <- NA
  year <- NA
  hs_product_code_2 <- NA
  hs_product_code_4 <- NA
  
  #--------Read and Confirm CountryCode input-----------------------------------
  
  CountryDataPath <- system.file("extdata", "country_codes_V202201.csv", package = "iopspackage")
  
  AllCountryCodes <- read.csv(CountryDataPath)
  countryChosen <- AllCountryCodes[AllCountryCodes$country_code == CountryCode, ]
  
  #Error conditions
  if( dim(countryChosen)[1] == 0) {
    stop('Invalid country code: No matches found')
  } else if( dim(countryChosen)[1] > 1) {
    message('More than one matching country found: ')
    n = 1
    for ( n in 1:dim(countryChosen)[1] ) {
      message(n, '. ',countryChosen$country_name_abbreviation[n])
    }
    CorrectCountryNr <- readline("Please enter correct country (corresponding integer): ")
    countryChosen <- countryChosen[CorrectCountryNr, ]
  }
  
  if( ComplexMethod == "reflections"&& ComplexMethod == "eigenvalues" && ComplexMethod == "fitness") stop('Invalid complexity method')
  
  #Determine country's 3-digit country code
  chosenCountry <- as.character(countryChosen["iso_3digit_alpha"])
  message(paste0("Selected country: ", chosenCountry))
  
  #--------Import Trade Data----------------------------------------------------
  
  if ( tradedigit == 4) {
    # Remove the import value data from the raw trade data
    tradeData <- select(tradeData, -c("j", "q"))
    # Aggregate export value at the 6-digit level
    tradeData <- tradeData %>%
      group_by(t, i, k) %>%
      summarize(v = sum(v), .groups = "keep")
    # Create the 4-digit codes
    digit_estimate <- tradeData['k']/100
    # Store codes as nearest lower integer
    digit_new <- floor(digit_estimate)
    # Add a new column to the export data with the 4-digit level code
    tradeData['k'] <- digit_new
    # Sum export values to 4-digit level and reset index
    tradeData <- tradeData %>%
      group_by(t, i, k) %>%
      summarize(v = sum(v))
    message("Export data 4-digit summation complete")
    
    if (is.null(GVCMapping) == FALSE) {
      #Update trade digit to 4 digits in the Value Chain provided
      colnames(GVCMapping) <- c("tierNumber", "GVCactivityNumber", "HScode")
      
      # Create the 4-digit codes
      digit_estimate_VC <- GVCMapping['HScode']/100
      # Store codes as nearest lower integer
      digit_new_VC <- floor(digit_estimate_VC)
      # Add a new column to the export data with the 4-digit level code
      GVCMapping['HScode'] <- digit_new_VC
      # Sum export values to 4-digit level and reset index
      GVCMapping <- GVCMapping %>%
        distinct(tierNumber, GVCactivityNumber, HScode)
      message("Value Chain HS 4-digit convertion complete")
    }
    
  } else {
    # Trade data transformation to retain exports at the 6-digit level
    # Remove import data from the raw trade data
    tradeData <- select(tradeData, -c("j", "q"))
    # Sum export value to 6-digit code level and reset the index
    tradeData <- tradeData %>%
      group_by(t, i, k) %>%
      summarize(v = sum(v), .groups = "keep")
    message("Export data 6-digit summation complete")
  }
  
  #++++++++++++Determine country's 3-digit country code+++++++++++++++++++++++++
  
  # read in trade data
  tradeData2018 <- tradeData #Possibly remove and work with tradeData in whole doc???
  tradeData2018 <- tradeData2018[c("t","v","i","k")]
  colnames(tradeData2018) <- c("year","export_value","country_code","hs_product_code")
  
  #++++++++++++Merge data sets (for compatibility)++++++++++++++++++++++++++++++
  
  ISO_Country <-AllCountryCodes[,c(1, 5)]
  colnames(ISO_Country)[2] <- "location_code"
  
  common_col_names <- intersect(names(tradeData2018), names(ISO_Country))
  
  MergedTradeData <- merge(tradeData2018, ISO_Country, by = common_col_names)
  
  #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  
  tradeData2018 <- MergedTradeData[,c("year","export_value","location_code","hs_product_code")]
  
  message("Trade data successfully imported")
  
  #-----------------------------------------------------------------------------
  
  expdata <- tradeData2018[ ,c(2, 3, 4)]
  wide_expdata <- expdata %>% pivot_wider(names_from = hs_product_code, values_from = export_value, values_fn = sum, names_sort = TRUE, values_fill = 0)  # transform the data so that all data for one country is contained in a single row
  
  #--------Calculate Mbin, Mabs, complexities, CSRCA----------------------------
  
  # create matrix of RCA values (Mabs) and binary RCA matrix (Mbin)
  Mabs <- balassa_index(expdata, discrete = FALSE, country = "location_code", product = "hs_product_code", value = "export_value")
  Mbin <- balassa_index(expdata, country = "location_code", product = "hs_product_code", value = "export_value")
  
  M_binary <- as.matrix(Mbin)
  M_absolute <- as.matrix(Mabs)
  
  message("M-matrices calculated")
  
  Products <- unique(expdata$hs_product_code)   #Roland: 1219 product codes, already sorted
  Countries <- rownames(Mabs)   #Roland: 235 countries, alphabetically ordered
  
  # create a matrix containing the RCA data for the country being analysed
  
  # chosen country subset: wide_expdata[which(wide_expdata$location_code == chosenCountry), -1] below
  CS_RCAmat <- cbind.data.frame(colnames(wide_expdata[which(wide_expdata$location_code == chosenCountry), -1]), as.numeric(wide_expdata[which(wide_expdata$location_code == chosenCountry), -1]), as.numeric(colSums(wide_expdata[ ,-1])), as.numeric(Mabs[chosenCountry, ]))
  colnames(CS_RCAmat) <- c("hs_product_code", paste(chosenCountry, "export_of_product", sep = "_"), "World_export_of_product", paste(chosenCountry, "RCA", sep = "_"))
  # || hs_product_code | ZAF_export_of_product | World_export_of_product | ZAF_RCA ||
  
  message("RCA calculated")
  
  #+++++++++COMPLEXITY MEASURES+++++++++++++++++++++++++++++++++++++++++++++++++
  
  #User selects one of three methods to calculate complexities,
  if (ComplexMethod == "eigenvalues"){
    
    message(paste0("Starting calculation of complexity measures using the '", ComplexMethod, "' method" ))
    
    # use binary M matrix to calculate complexity using EIGENVALUES, number of iterations must be provided
    #NOTE: This specific method takes a significantly long time to compute
    Compl_Meth <- complexity_measures(Mbin, method = "eigenvalues")
    
    Method <- "eigv" #method descriptor for Excel/CSV file
    
  } else if(ComplexMethod == "reflections") {
    
    message(paste0("Starting calculation of complexity measures using the '", ComplexMethod, "' method" ))
    
    # use binary M matrix to calculate complexity using METHOD OF REFLECTIONS, number of iterations must be provided
    Compl_Meth <- complexity_measures(Mbin, method = "reflections", iterations = iterCompl)
    
    Method <- "refl"
    
  } else {
    
    message(paste0("Starting calculation of complexity measures using the '", ComplexMethod, "' method" ))
    
    # use binary M matrix to calculate complexity using FITNESS COMPLEXITY, number of iterations must be provided
    Compl_Meth <- complexity_measures(Mbin, method = "fitness", iterations = iterCompl)
    
    Method <- "fitn"
  }
  
  #determine complexities with input method
  ECI <- Compl_Meth$complexity_index_country
  PCI <- Compl_Meth$complexity_index_product
  
  #--------Calculate proximity, opp gain, opp val, distance, density------------
  
  # calculate the proximity between products
  proxes <- proximity(Mbin)
  
  Proximities <- proxes$proximity_product
  Similarities <- proxes$proximity_country
  
  # compute the opportunity value
  opporVal <- complexity_outlook(Mbin, Proximities, PCI)
  cntryOpporVal <- opporVal$complexity_outlook_index
  opporGain <- opporVal$complexity_outlook_gain
  Opportunity_Gain <- as.matrix(opporGain)
  
  # compute the densities and distances
  # Note: (1 - Mbin) creates matrix of ones where Mbin has zeroes and vice versa
  distance <- tcrossprod(1 - as.matrix(Mbin), as.matrix(Proximities)/rowSums(as.matrix(Proximities)))
  density <- tcrossprod(as.matrix(Mbin), as.matrix(Proximities)/rowSums(as.matrix(Proximities)))
  
  message("Complexity measures calculated")
  
  #--------Create GVCfull, GVCacts, GVCtiers------------------------------------
  
  #--------Read in industry value chain mapping---------------------------------
  
  if (is.null(GVCMapping) == TRUE) {
    # Store the subset of data associated with the input country code.
    merged_country_df <- subset(tradeData2018, location_code == chosenCountry) #Possibly do with code and not ISO 3 name?? That's how its done in python
    message("Pre-mappings complete")
    
    # Binary RCA assignment for IO-PS calculations.
    
    #Add the rca column
    merged_country_df <- merge(merged_country_df, CS_RCAmat[, c("hs_product_code", "ZAF_RCA")], by = "hs_product_code", all.x = TRUE)
    # Rename the ZAF_RCA column to rca
    colnames(merged_country_df)[colnames(merged_country_df) == "ZAF_RCA"] <- "rca"
    
    #Add the distance column
    # Convert original distance matrix to long format with hs_product_code, location_code, and distance columns
    distance_long <- as.data.frame(as.table(distance))
    colnames(distance_long) <- c("location_code", "hs_product_code", "distance")
    # Merge the distance data into merged_country_df based on location_code and hs_product_code
    merged_country_df <- merge(merged_country_df, distance_long, by = c("location_code", "hs_product_code"), all.x = TRUE)
    
    #Add the PCI column
    #Convert PCI list into a table with product code and corresponding pci values
    PCI_formated <- as.data.frame(as.table(Compl_Meth$complexity_index_product))
    colnames(PCI_formated) <- c("hs_product_code", "PCI")
    merged_country_df <- merge(merged_country_df, PCI_formated, by = "hs_product_code", all.x = TRUE)
    
    #Add Mcp column
    #Create column named Mcp and asssign 0s
    merged_country_df$Mcp <- 0
    # Set Mcp column to a value of 1 for all products in the value chain for which the country has an RCA > 1.
    merged_country_df$Mcp[merged_country_df$rca >= 1] <- 1
    
    message("No Value Chain input provided, calculating default at all trade digit levels") #Change wording??
    
    # Aggregate raw trade data to the user specified trade digit level.
    if (tradedigit == 4) {
      message('No value chain input, calculating default for 4-digit input')
      # Write 4-digit aggregation
      fourcountry_df <- merged_country_df[, c("year", "hs_product_code", "Mcp", "rca", "distance", "PCI")]
      # Store column headings
      colnames(fourcountry_df) <- c('year', 'HS Product code', 'Mcp', 'Average RCA', 'Average distance', 'Average product complexity')
      
      message('No value chain 4-digit results calculated')
      
      message('No value chain input, calculating default for 2-digit input')
      # Create 2-digit level codes
      digit_estimate <- merged_country_df['hs_product_code'] / 100
      # Store codes as nearest lower integer
      digit_new <- floor(digit_estimate)
      # Add new column at 2-digit level
      merged_country_df['hs_product_code_2'] <- digit_new
      # Store new digits and key metrics
      twocountry_df <- merged_country_df[, c('year', 'hs_product_code_2', 'rca', 'distance', 'PCI')]
      # Sum export values to new 2-digit level and reset index
      twocountry_df <- twocountry_df %>%
        group_by(year, hs_product_code_2) %>%
        summarize_all(mean) %>%
        ungroup()
      # Store column headings
      colnames(twocountry_df) <- c('year', 'HS Product code', 'Average RCA', 'Average distance', 'Average product complexity')
      #Remove hs_product_code_2 column
      merged_country_df <- select(merged_country_df, -c("hs_product_code_2"))
      
      message('No value chain 2-digit results calculated')
      
      #Write results to.xlsx and .csv file
      message("Writing results to a .xlsx file")
      # Create a new workbook
      wb <- createWorkbook()
      
      # Add the first dataset to the workbook on a new sheet named "noGVC4Digit_Results"
      addWorksheet(wb, "noGVC4Digit_Results")
      writeData(wb, "noGVC4Digit_Results", fourcountry_df)
      
      # Add the second dataset to the workbook on a new sheet named "noGVC2Digit_Results"
      addWorksheet(wb, "noGVC2Digit_Results")
      writeData(wb, "noGVC2Digit_Results", twocountry_df)
      
      # Save the workbook to a file
      saveWorkbook(wb, "Combined_Results_NoVC.xlsx", overwrite = TRUE)
      
      
      message("Writing results to a .csv file")
      # Create a new workbook
      wb <- createWorkbook()
      
      # Add the first dataset to the workbook on a new sheet named "noGVC4Digit_Results"
      addWorksheet(wb, "noGVC4Digit_Results")
      writeData(wb, "noGVC4Digit_Results", fourcountry_df)
      
      # Add the second dataset to the workbook on a new sheet named "noGVC2Digit_Results"
      addWorksheet(wb, "noGVC2Digit_Results")
      writeData(wb, "noGVC2Digit_Results", twocountry_df)
      
      # Save the workbook to a file
      saveWorkbook(wb, "Combined_Results_NoVC.csv", overwrite = TRUE)
      
      
    } else {
      message('No value chain input, calculating default for 6-digit input')
      # Store 6-digit aggregation data
      sixcountry_df <- merged_country_df[, c("year", "hs_product_code", "Mcp", "rca", "distance", "PCI")]
      # Store column headings
      colnames(sixcountry_df) <- c('year', 'HS Product code', 'Mcp', 'Average RCA', 'Average distance', 'Average product complexity')
      
      message('No value chain 6-digit results calculated')
      
      message('No value chain input, calculating default for 4-digit input')
      # Create 4-digit codes
      digit_estimate <- merged_country_df['hs_product_code'] / 100
      # Store 4-digit estimates as nearest lower integer
      digit_new <- floor(digit_estimate)
      # Add a new column containing the 4-digit level codes
      merged_country_df['hs_product_code_4'] <- digit_new
      # Store new digits and key metrics
      fourcountry_df <- merged_country_df[, c('year', 'hs_product_code_4', 'rca', 'distance', 'PCI')]
      # Sum export values to new 4-digit level and reset index
      fourcountry_df <- fourcountry_df %>%
        group_by(year, hs_product_code_4) %>%
        summarize_all(mean) %>%
        ungroup()
      # Store column headings
      colnames(fourcountry_df) <- c('year', 'HS Product code', 'Average RCA', 'Average distance', 'Average product complexity')
      
      message('No value chain 4-digit results calculated')
      
      message('No value chain input, calculating default for 2-digit input')
      # Create 2-digit level codes
      digit_estimate <- merged_country_df['hs_product_code'] / 10000
      # Store codes as nearest lower integer
      digit_new <- floor(digit_estimate)
      # Add new column at 2-digit level
      merged_country_df['hs_product_code_2'] <- digit_new
      # Store new digits and key metrics
      twocountry_df <- merged_country_df[, c('year', 'hs_product_code_2', 'rca', 'distance', 'PCI')]
      # Sum export values to new 2-digit level and reset index
      twocountry_df <- twocountry_df %>%
        group_by(year, hs_product_code_2) %>%
        summarize_all(mean) %>%
        ungroup()
      # Store column headings
      colnames(twocountry_df) <- c('year', 'HS Product code', 'Average RCA', 'Average distance', 'Average product complexity')
      
      message('No value chain 2-digit results calculated')
      
      
      #Write results to.xlsx and .csv file
      message("Writing results to a .xlsx file")
      # Create a new workbook
      wb <- createWorkbook()
      
      # Add the first dataset to the workbook on a new sheet named "noGVC6Digit_Results"
      addWorksheet(wb, "noGVC6Digit_Results")
      writeData(wb, "noGVC6Digit_Results", sixcountry_df)
      
      # Add the second dataset to the workbook on a new sheet named "noGVC4Digit_Results"
      addWorksheet(wb, "noGVC4Digit_Results")
      writeData(wb, "noGVC4Digit_Results", fourcountry_df)
      
      # Add the third dataset to the workbook on a new sheet named "noGVC2Digit_Results"
      addWorksheet(wb, "noGVC2Digit_Results")
      writeData(wb, "noGVC2Digit_Results", twocountry_df)
      
      # Save the workbook to a file
      saveWorkbook(wb, "Combined_Results_NoVC.xlsx", overwrite = TRUE)
      
      message("Writing results to a .csv file")
      # Create a new workbook
      wb <- createWorkbook()
      
      # Add the first dataset to the workbook on a new sheet named "noGVC6Digit_Results"
      addWorksheet(wb, "noGVC6Digit_Results")
      writeData(wb, "noGVC6Digit_Results", sixcountry_df)
      
      # Add the second dataset to the workbook on a new sheet named "noGVC4Digit_Results"
      addWorksheet(wb, "noGVC4Digit_Results")
      writeData(wb, "noGVC4Digit_Results", fourcountry_df)
      
      # Add the third dataset to the workbook on a new sheet named "noGVC2Digit_Results"
      addWorksheet(wb, "noGVC2Digit_Results")
      writeData(wb, "noGVC2Digit_Results", twocountry_df)
      
      # Save the workbook to a file
      saveWorkbook(wb, "Combined_Results_NoVC.csv", overwrite = TRUE)
    }
  } else {
    #GVCMapping: || tierNumber | GVCactivityNumber | HScode ||
    colnames(GVCMapping) <- c("tierNumber", "GVCactivityNumber", "HScode")
    
    message("Value chain succesfully imported")
    
    message("Starting IO-PS calculations")
    
    GVCMapping <- GVCMapping[order(GVCMapping$tierNumber, GVCMapping$GVCactivityNumber, GVCMapping$HScode), ]   # reorder GVCMapping according to tier, then activity, then HScode
    
    numTiers <- length(unique(GVCMapping$tierNumber))         #Roland: = 5
    numActs <- length(unique(GVCMapping$GVCactivityNumber))   #Roland: = 49
    numProds <- length(unique(GVCMapping$HScode))             #Roland: = 272
    
    numProdsInAct <- as.matrix(GVCMapping[ ,c(2, 3)] %>% group_by(GVCactivityNumber) %>% tally())[ ,2]
    numActsInTier <- as.matrix(unique(GVCMapping[ ,c(1, 2)]) %>% group_by(tierNumber) %>% tally())[ ,2]
    numProdsInTier <- as.matrix(GVCMapping[ ,c(1, 3)] %>% group_by(tierNumber) %>% tally())[ ,2]
    
    
    #-----GVCfull-----
    GVCfull <- as.data.frame(matrix(rep(0, times = (numProds)*15), nrow = numProds, ncol = 15))   # initialise GVCfull
    colnames(GVCfull) <- c("tierNumber", "GVCactivityNumber", "HScode", "Complexity", "Density", "RCA", "Complexity_ifOpp", "Distance_ifOpp", "OpporGain_ifOpp", "RCA_ifOpp", paste(chosenCountry, "export_of_product", sep = "_"), "World_export_of_product", paste(chosenCountry, "export_of_product_ifOpp", sep = "_"), "World_export_of_product_ifOpp", "productPosition")
    
    GVCfull[ , c(1, 2, 3)] <- GVCMapping   #Populate first 3 columns with GVCMapping values
    GVCfull <- GVCfull[order(GVCfull$HScode), ]   # order according to HS code
    
    GVCfull[ , 4] <- as.numeric(PCI[which(Products %in% GVCMapping$HScode)])    #Populate 4th column with PCI values
    GVCfull[ , 5] <- as.numeric(density[chosenCountry, which(Products %in% GVCMapping$HScode)])   #Populate 5th column with PCI values. For distances and densities, subset the one row that is specific to the country being analysed
    GVCfull[ , 6] <- CS_RCAmat[which(Products %in% GVCMapping$HScode), 4]    #Populate 6th column with RCA values
    GVCfull[ , c(11, 12)] <- CS_RCAmat[which(Products %in% GVCMapping$HScode), c(2, 3)] #Populate 11th and 12th column with RCA values
    GVCfull[ , 15] <- which(Products %in% GVCMapping$HScode)  #Populate 15th column with Products values
    
    GVCfull[which(GVCfull$RCA < 1), c(7, 10, 13, 14)] <- GVCfull[which(GVCfull$RCA < 1), c(4, 6, 11, 12)] #Populate 7th, 10th, 13th, 14th columns that has advantage?
    
    GVCfull[ , 8] <- as.numeric(distance[chosenCountry, which(Products %in% GVCMapping$HScode)])  #Populate 8th column with distance values
    GVCfull[-which(GVCfull$RCA < 1), 8] <- 0
    
    GVCfull[ , 9] <- as.numeric(opporGain[chosenCountry, which(Products %in% GVCMapping$HScode)])  #Populate 9th column with opporGain values
    GVCfull[-which(GVCfull$RCA < 1), 9] <- 0
    
    GVCfull <- GVCfull[order(GVCfull$tierNumber, GVCfull$GVCactivityNumber, GVCfull$HScode), ]
    
    #-----GVCacts-----
    numOppProdsInAct <- rep(0, times = length(numProdsInAct))
    
    numOppProdsInAct[as.matrix(GVCMapping[which(GVCMapping$HScode %in% GVCfull[which(GVCfull$RCA < 1), 3]), ] %>% group_by(GVCactivityNumber) %>% tally())[ ,1]] <- as.matrix(GVCMapping[which(GVCMapping$HScode %in% GVCfull[which(GVCfull$RCA < 1), 3]), ] %>% group_by(GVCactivityNumber) %>% tally())[ ,2]
    
    #in case of mismatch between data lengths
    if (length(numProdsInAct)<length(numOppProdsInAct)) {
      stop("Error: data selected for 'tradeData' and 'GVCmapping' isn't compatible with each other. Select data within the same 'H' class, i.e., 'H0', 'H3', or 'H5'")
    }
    
    GVCacts <- as.data.frame(matrix(rep(0, times = numActs*15), nrow = numActs, ncol = 15))   # initialise GVCacts
    colnames(GVCacts) <- c("tierNumber", "GVCactivityNumber", "AvgComplexity", "AvgDensity", "AvgRCA", "AvgComplexity_ifOpp", "AvgDistance_ifOpp", "AvgOpporGain_ifOpp", "AvgRCA_ifOpp", paste(chosenCountry, "export_of_activity", sep = "_"), "World_export_of_activity", paste(chosenCountry, "export_of_activity_ifOpp", sep = "_"), "World_export_of_activity_ifOpp", "NumberOfProductsInActivity", "NumberOfOppProductsInActivity")
    
    GVCacts[ , c(1, 2)] <- unique(GVCMapping[ ,c(1, 2)])  #Populate 1st and 2nd columns with GVCMapping values
    
    GVCacts[ , c(3:5)] <- aggregate(GVCfull[ , c(4:6)], by = list(activity = GVCfull$GVCactivityNumber), FUN = sum)[, -1]/numProdsInAct  #Populate 3th, 4th and 5th columns with average values   # also works: #aggregate(GVCfull[ , c(4:6)], by = list(activity = GVCfull$GVCactivityNumber), FUN = mean)[ , -1]
    #GVCacts[-which(numOppProdsInAct == 0) , c(6:9)] <- aggregate(GVCfull[-which(GVCfull$GVCactivityNumber %in% which(numOppProdsInAct == 0)) , c(7:10)], by = list(activity = GVCfull$GVCactivityNumber[-which(GVCfull$GVCactivityNumber %in% which(numOppProdsInAct == 0))]), FUN = sum)[, -1]/(numOppProdsInAct[-which(numOppProdsInAct == 0)])  #Populate 6th, 7th, 8th columns
    
    #in case of no values numOppProdsInAct == 0, the first line then breaks, hence the else statement
    if(sum(GVCfull$GVCactivityNumber %in% which(numOppProdsInAct == 0) == TRUE) > 0){
      GVCacts[-which(numOppProdsInAct == 0) , c(6:9)] <- aggregate(GVCfull[-which(GVCfull$GVCactivityNumber %in% which(numOppProdsInAct == 0)) , c(7:10)], by = list(activity = GVCfull$GVCactivityNumber[-which(GVCfull$GVCactivityNumber %in% which(numOppProdsInAct == 0))]), FUN = sum)[, -1]/(numOppProdsInAct[-which(numOppProdsInAct == 0)])
    } else {
      GVCacts[, c(6:9)] <- aggregate(GVCfull[ ,c(7:10)], by = list(activity = GVCfull$GVCactivityNumber), FUN = mean)[, -1]
      J <-"Else"
    }
    
    GVCacts[ , c(10:13)] <- aggregate(GVCfull[ , c(11:14)], by = list(activity = GVCfull$GVCactivityNumber), FUN = sum)[ , -1]  #Populate 10th, 11th, 12th columns with  values from 11, 12 and 13th columns
    GVCacts[ , 14] <- numProdsInAct  #Populate 14th column with numProdsInAct values
    GVCacts[ , 15] <- numOppProdsInAct  #Populate 15th column with numOppProdsInAct values
    
    
    #-----GVCtiers-----
    
    numOppProdsInTier <- rep(0, times = length(numActsInTier))
    numOppActsInTier <- rep(0, times = length(numActsInTier))
    
    numOppProdsInTier[as.matrix(GVCMapping[which(GVCMapping$HScode %in% GVCfull[which(GVCfull$RCA < 1), 3]), ] %>% group_by(tierNumber) %>% tally())[ ,1]] <- as.matrix(GVCMapping[which(GVCMapping$HScode %in% GVCfull[which(GVCfull$RCA < 1), 3]), ] %>% group_by(tierNumber) %>% tally())[ ,2]
    numOppActsInTier[as.matrix(GVCacts[which(GVCacts$GVCactivityNumber %in% GVCacts[which(GVCacts$AvgRCA < 1), 2]), ] %>% group_by(tierNumber) %>% tally())[ ,1]] <- as.matrix(GVCacts[which(GVCacts$GVCactivityNumber %in% GVCacts[which(GVCacts$AvgRCA < 1), 2]), ] %>% group_by(tierNumber) %>% tally())[ ,2]
    
    GVCtiers <- as.data.frame(matrix(rep(0, times = numTiers*14), nrow = numTiers, ncol = 14))   # initialise GVCtiers
    colnames(GVCtiers) <- c("tierNumber", "AvgComplexity", "AvgDensity", "AvgRCA", "AvgComplexity_ifOpp", "AvgDistance_ifOpp", "AvgOpporGain_ifOpp", "AvgRCA_ifOpp", paste(chosenCountry, "export", sep = "_"), "World_export", paste(chosenCountry, "export_ifOpp", sep = "_"), "World_export_ifOpp", "NumberOfActivitiesInTier", "NumberOfOppActivitiesInTier")
    
    #in case of no values numOppProdsInTier == 0, the first line then breaks, hence the else statement
    if(sum(GVCfull$tierNumber %in% which(numOppProdsInTier == 0) == TRUE) > 0){
      GVCtiers[-which(numOppProdsInTier == 0) , c(5:8)] <- aggregate(GVCfull[-which(GVCfull$tierNumber %in% which(numOppProdsInTier == 0)) , c(7:10)], by = list(tier = GVCfull$tierNumber[-which(GVCfull$tierNumber %in% which(numOppProdsInTier == 0))]), FUN = sum)[, -1]/(numOppProdsInTier[-which(numOppProdsInTier == 0)])
    } else {
      GVCtiers[, c(5:8)] <- aggregate(GVCfull[ ,c(7:10)], by = list(tier = GVCfull$tierNumber), FUN = mean)[, -1]
      J <-"Else"
    }
    
    GVCtiers[ , 1] <- unique(GVCMapping$tierNumber)  #Populate 1st column with tiernumber values
    GVCtiers[ , c(2:4)] <- aggregate(GVCfull[ , c(4:6)], by = list(tier = GVCfull$tierNumber), FUN = sum)[, -1]/numProdsInTier  #Populate 2nd, 3rd and 4th column with  values
    GVCtiers[ , c(9:12)] <- aggregate(GVCfull[ , c(11:14)], by = list(tier = GVCfull$tierNumber), FUN = sum)[ , -1]  #Populate 9th, 10th, 11th and 12th columns with values from column 11, 12, 13 and 14.
    GVCtiers[ , 13] <- numActsInTier  #Populate 13th column with numActsInTier values
    GVCtiers[ , 14] <- numOppActsInTier  #Populate 13th column with numOppActsInTier values
    
    #/////////////////////Results to be Exported////////////////////////////////
    
    Product_Results <- GVCfull
    Product_Category_Results <- GVCacts
    Tier_Results <- GVCtiers
  }
  #--------Converting back to 'country_codes'-----------------------------------
  
  #ECI
  ECI_df <- as.data.frame(ECI, check.names=FALSE)  #Initialise wrong ECI values
  ECI_df <- cbind(rownames(ECI_df), data.frame(ECI_df, row.names=NULL))  #Bind country names into data frame (doesn't match correct ECI value)
  colnames(ECI_df)[1] <- "location_code"
  ECI_df <- merge.data.frame(ISO_Country, ECI_df, by = intersect(names(ISO_Country), names(ECI_df)))[, c(2,3)]  #Match and move correct country with correct ECI value
  ECI_df <- data.frame(ECI_df[order(ECI_df$country_code),],  row.names=1) #Remove country name and reorder accourding to codes
  
  #PCI
  PCI_df <- as.data.frame(PCI) #Initialise PCI values
  
  #Opportunity Gain
  OG_df <- as.data.frame(Opportunity_Gain)  #Initialise OG values
  OG_df <- cbind(rownames(OG_df), data.frame(OG_df, row.names=NULL))  #Bind country names into data frame (doesn't match correct OG values)
  colnames(OG_df)[1] <- "location_code"
  OG_df <- merge.data.frame(ISO_Country, OG_df, by = intersect(names(ISO_Country), names(OG_df)))[,c(-1)]  #Replace Country 3-digit name with corresponding Country code, matching to OG values
  OG_df <- data.frame(OG_df[order(OG_df$country_code),],  row.names=1)  #Reorder and remove country codes
  names(OG_df) <- sub("^X", "", names(OG_df))  #Remove X from names
  
  #distance
  dist_df <- as.data.frame(distance)  #Initialise distance values
  dist_df <- cbind(rownames(dist_df), data.frame(dist_df, row.names=NULL))  #Bind country names into data frame (doesn't match correct distance values)
  colnames(dist_df)[1] <- "location_code"
  dist_df <- merge.data.frame(ISO_Country, dist_df, by = intersect(names(ISO_Country), names(dist_df)))[,c(-1)]  #Replace Country 3-digit name with corresponding Country code, matching to distance values
  dist_df <- data.frame(dist_df[order(dist_df$country_code),],  row.names=1)  #Reorder and remove country codes
  names(dist_df) <- sub("^X", "", names(dist_df))  #Remove X from names
  dist_df <- as.matrix(dist_df)  #Convert from data frame to matrix
  
  #density
  dens_df <- as.data.frame(density)  #Initialise density values
  dens_df <- cbind(rownames(dens_df), data.frame(dens_df, row.names=NULL))  #Bind country names into data frame (doesn't match correct distance values)
  colnames(dens_df)[1] <- "location_code"
  dens_df <- merge.data.frame(ISO_Country, dens_df, by = intersect(names(ISO_Country), names(dens_df)))[,c(-1)]  #Replace Country 3-digit name with corresponding Country code, matching to density values
  dens_df <- data.frame(dens_df[order(dens_df$country_code),],  row.names=1)  #Reorder and remove country codes
  names(dens_df) <- sub("^X", "", names(dens_df))  #Remove X from names
  dens_df <- as.matrix(dens_df)  #Convert from data frame to matrix
  
  #M-absolute
  Mabs_df <- as.data.frame(M_absolute)  #Initialise Mabs values
  Mabs_df <- cbind(rownames(Mabs_df), data.frame(Mabs_df, row.names=NULL))  #Bind country names into data frame (doesn't match correct Mabs values)
  colnames(Mabs_df)[1] <- "location_code"
  Mabs_df <- merge.data.frame(ISO_Country, Mabs_df, by = intersect(names(ISO_Country), names(Mabs_df)))[,c(-1)]  #Replace Country 3-digit name with corresponding Country code, matching to Mabs values
  Mabs_df <- data.frame(Mabs_df[order(Mabs_df$country_code),],  row.names=1)  #Reorder and remove country codes
  names(Mabs_df) <- sub("^X", "", names(Mabs_df))  #Remove X from names
  Mabs_df <- as.matrix(Mabs_df)  #Convert from data frame to matrix
  
  #M-binary
  Mbin_df <- as.data.frame(M_binary)  #Initialise Mbin values
  Mbin_df <- cbind(rownames(Mbin_df), data.frame(Mbin_df, row.names=NULL))  #Bind country names into data frame (doesn't match correct Mbin values)
  colnames(Mbin_df)[1] <- "location_code"
  Mbin_df <- merge.data.frame(ISO_Country, Mbin_df, by = intersect(names(ISO_Country), names(Mbin_df)))[,c(-1)]  #Replace Country 3-digit name with corresponding Country code, matching to Mbin values
  Mbin_df <- data.frame(Mbin_df[order(Mbin_df$country_code),],  row.names=1)  #Reorder and remove country codes
  names(Mbin_df) <- sub("^X", "", names(Mbin_df))  #Remove X from names
  Mbin_df <- as.matrix(Mabs_df)  #Convert from data frame to matrix
  #-----------------------------------------------------------------------------
  
  if (is.null(GVCMapping) == TRUE){
    ReturnIOPS <- list(ECI_df, PCI_df, OG_df, dist_df, dens_df, Mabs_df, Mbin_df) #Possibly add twocountry_df etc?? Also check write format?
    names(ReturnIOPS) <- c("ECI", "PCI", "Opportunity_Gain", "distance", "density", "M_absolute", "M_binary")
  } else {
    ReturnIOPS <- list(ECI_df, PCI_df, OG_df, dist_df, dens_df, Mabs_df, Mbin_df, Tier_Results, Product_Category_Results, Product_Results)
    names(ReturnIOPS) <- c("ECI", "PCI", "Opportunity_Gain", "distance", "density", "M_absolute", "M_binary", "Tier_Results", "Product_Category_Results", "Product_Results")
    
    #------------Write to Excel and CSV files---------------
    message("Writing results to a .xlsx file")
    # Create a new workbook
    wb <- createWorkbook()
    # Add the first dataset to the workbook on a new sheet named "Product_Results"
    addWorksheet(wb, "Product_Results")
    writeData(wb, "Product_Results", ReturnIOPS$Product_Results)
    # Add the second dataset to the workbook on a new sheet named "Prod_Cat_Results"
    addWorksheet(wb, "Prod_Cat_Results")
    writeData(wb, "Prod_Cat_Results", ReturnIOPS$Product_Category_Results)
    # Add the third dataset to the workbook on a new sheet named "Tier_Results"
    addWorksheet(wb, "Tier_Results")
    writeData(wb, "Tier_Results", ReturnIOPS$Tier_Results)
    # Save the workbook to a file
    saveWorkbook(wb, "Combined_Results.xlsx", overwrite = TRUE)
    
    message("Writing results to a .csv file")
    # Create a new workbook
    wb <- createWorkbook()
    # Add the first dataset to the workbook on a new sheet named "Product_Results"
    addWorksheet(wb, "Product_Results")
    writeData(wb, "Product_Results", ReturnIOPS$Product_Results)
    # Add the second dataset to the workbook on a new sheet named "Prod_Cat_Results"
    addWorksheet(wb, "Prod_Cat_Results")
    writeData(wb, "Prod_Cat_Results", ReturnIOPS$Product_Category_Results)
    # Add the third dataset to the workbook on a new sheet named "Tier_Results"
    addWorksheet(wb, "Tier_Results")
    writeData(wb, "Tier_Results", ReturnIOPS$Tier_Results)
    # Save the workbook to a file
    saveWorkbook(wb, "Combined_Results.csv", overwrite = TRUE)
    #-------------------------------------------------------
  }
  message("--------DONE!--------")
  return(ReturnIOPS)
}
