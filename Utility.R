# install.packages("\\\\bankspb.ru\\common\\Terminal Internet\\reuters\\Downloads\\RODBC_1.3-14.zip",repos = NULL, type="source")
# install.packages("\\\\bankspb.ru\\common\\Terminal Internet\\reuters\\Downloads\\Rcpp_0.12.7.zip",repos = NULL, type="source")
# install.packages("\\\\bankspb.ru\\common\\Terminal Internet\\reuters\\Downloads\\readxl_0.1.1.zip",repos = NULL, type="source")

#RShowDoc("RODBC", package="RODBC")
library(RODBC)
library(readxl)

driver = "DRIVER=Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)"
identityColumn <- "Names" #support column name

#recieve types of risk factors
getFactors <- function(directory,fileName){
  cat(" Geting Factors... \t")
  fileNameNormalized=normalizePath(file.path(directory,fileName))
  folder = paste("DBQ=",fileNameNormalized)
  constr = paste(driver, folder, "readOnly=FALSE", sep=";")
  channel <- odbcDriverConnect(constr)
  types=data.frame(matrix(0,ncol = 0, nrow = 0),stringsAsFactors=F)
  types <- sqlQuery(channel, "select * from [Factors$]")
  colnames(types) <- c(identityColumn, "Group", "Type") #rename columns
  cat(" Factors done \n")
  return(types)
  # odbcClose(channel)
  odbcCloseAll()
}

#function to get covariance matrix with first column as Names of factors
#RODBC not working properly here
getCovar <- function(directory,fileName) {
  odbcCloseAll()
  cat(" Geting Covariance... \t")
  fileNameNormalized=normalizePath(file.path(directory,fileName))
  covar = data.frame(matrix(0,ncol = 0, nrow = 0),stringsAsFactors=F)
  covar <-  read_excel(fileNameNormalized, sheet = "Covar", col_names = T, col_types = NULL, na = "", skip = 0)
  colnames(covar)[1] <- identityColumn
  cat(" Covariance done \n")
  return(covar)
}

#get from all factors only useful
truncFactors <- function(allFactors, options, covar){
  if (is.null(options) || is.na(options) || length(options) <= 0){
    cat("WARNING! Wrong options, Factors are unchanged")
    return(allFactors)
  }
  cat(" Truncing Factors... \t")
  factorNames <- unique(tolower(covar[,identityColumn])) #there are same factors in covariance matrix and with different registers
  factors <- data.frame(matrix(0, ncol = 1, nrow = length(factorNames), 
                               dimnames = list(NULL, c(identityColumn))),stringsAsFactors=F)
  factors[,identityColumn] <- factorNames
  riskfactors <- allFactors[which(allFactors[,"Group"] %in% options),] #trunc by options
  riskfactors[,identityColumn] <- tolower(riskfactors[,identityColumn]) #stundartise names
  factors <- merge(factors, riskfactors, by = identityColumn, all = F, sort = F)
  cat(" Factors trunced \n")
  return(factors)
}

#get from full covariance matrix only useful values and without support column
truncCovar <- function(fullCovar, truncFactors){
  if (is.null(truncFactors) || is.na(truncFactors) || length(truncFactors) <= 0){
    cat("WARNING! Wrong factors, covar matrix unchanged")
    return(fullCovar)
  }
  cat(" Truncing Covariance... \t")
  fullCovar <- fullCovar[,-c(1)] #remove first column with names
  names <- truncFactors[,identityColumn]
  columns <- unique(tolower(colnames(fullCovar))) #there are same factors in covariance matrix
  indexes <- which(columns %in% names)
  covar <- fullCovar[indexes, indexes]
  colnames(covar) <- tolower(colnames(covar))
  rownames(covar) <- colnames(covar)
  cat(" Covariance trunced \n")
  return(covar)
}

#function to get all sensitivities
getSensitivity <- function(directory, fileName, covar, factors){
  cat(" Geting sensitivities... \t")
  factorNames <- factors[,identityColumn]
  fileNameNormalized=normalizePath(file.path(directory,fileName))
  folder = paste("DBQ=",fileNameNormalized)
  constr = paste(driver, folder, "readOnly=FALSE", sep=";")
  channel <- odbcDriverConnect(constr)
  #clear sensitivities
  factors[,c("Sens1","Sens2")] <- 0
  rownames(factors) <- factorNames
  #collect MMsen sens
  MMsen <- data.frame(matrix(0,ncol = 0, nrow = 0),stringsAsFactors=F)
  MMsen <- sqlQuery(channel, "select * from [MMsen$]")
  MMsen <- MMsen[-c(1),]#delete first row
  MMsens1 <- MMsen[complete.cases(MMsen[,5]), c(4,5)] #remove all columns except 4:5 and rows with empty 5th column value
  colnames(MMsens1) <- c(identityColumn, " ") #rename columns
  factors <- uniteSensitivity(MMsens1, factors, "Sens1")
  #collect FXsen sens
  FXsen <- data.frame(matrix(0,ncol = 0, nrow = 0),stringsAsFactors=F)
  FXsen <- sqlQuery(channel, "select * from [FXsen$]")
  FXsens1 <- FXsen[28:50,1:11] #get this table as sens1 for FX
  colnames(FXsens1)[1] <- c(identityColumn)
  factors <- uniteSensitivity(FXsens1, factors, "Sens1")
  #collect BNsen sens
  BNsen <- data.frame(matrix(0,ncol = 0, nrow = 0),stringsAsFactors=F)
  BNsen <- sqlQuery(channel, "select * from [BNsen$]")
  BNsens1 <- BNsen[1:28,c(1,4:8)] #get this table as sens1 for BN
  colnames(BNsens1)[1] <- c(identityColumn) #names column
  colnames(BNsens1)[6] <- " " #futures column unnamed
  factors <- uniteSensitivity(BNsens1, factors, "Sens1")
  cat(" All sensitivities done \n")
  return(factors)
  # odbcClose(channel)
  odbcCloseAll()
}


#function that apply sensitivity from table to data frame in toColumnName column
#should change to rownames
uniteSensitivity <- function(from, to, toColumnName){
  for (i in from[,identityColumn]){
    if (!is.na(i) && length(i)>0){
      for (j in 2:length(colnames(from))){
        value <- from[which(i==from[,identityColumn]),j]
        if (!is.na(value) && is.double(value)){
          pat <- paste0("^",tolower(i),".?",tolower(colnames(from)[j]),"$")#pattern to paste sens to name started with i, and ending with j  
          risk <- grep(pat, rownames(to),value = T)#get row of this risk
          to[risk,toColumnName] <- value
        }
      }
    }
  }
  cat("#")
  return(to)
}

#get constant by options
getConstants <- function(directory, fileName, options, size){
  if (is.null(options) || is.na(options) || length(options) <= 0){
    cat(" WARNING! Wrong options, can't get constants ")
    return(0)
  }
  cat(" Geting Constants... \t")
  Constants <- matrix(0, nrow = length(options), ncol = size, dimnames = list(options, NULL))
  fileNameNormalized=normalizePath(file.path(directory,fileName))
  folder = paste("DBQ=",fileNameNormalized)
  constr = paste(driver, folder, "readOnly=FALSE", sep=";")
  channel <- odbcDriverConnect(constr)
  if ("RB" %in% options){
    BNsen <- data.frame(matrix(0,ncol = 0, nrow = 0),stringsAsFactors=F)
    BNsen <- sqlQuery(channel, "select * from [BNsen$]")
    BNconst <- BNsen[21:22,c(1,8)] #get this table as const for BN
    colnames(BNconst)[1] <- c(identityColumn)
    value <- BNconst[which(BNconst[,identityColumn]=="CorpCCC"), 2]
    if (!is.na(value) && is.double(value)){
      Constants[c("RB"),] <- value*0.004
    }
  }
  if ("FX" %in% options){
    FXsen <- data.frame(matrix(0,ncol = 0, nrow = 0),stringsAsFactors=F)
    FXsen <- sqlQuery(channel, "select * from [FXsen$]")
    FXconst <- FXsen[2:11,c(21,24)] #get this table as const for FX
    colnames(FXconst)[1] <- c(identityColumn)
    value <- sum(FXconst[, 2])
    if (!is.na(value) && is.double(value)){
      Constants["FX",] <- value
    }
  }
  cat(" Constants done \n")
  return(Constants)
  # odbcClose(channel)
  odbcCloseAll()
}

#generate realization matrix with size simulations with replacement eigenValues<0 to 0
simulate <- function(size, factors, eigen){
  cat(" Generating Simulations... \t")
  values <- eigen$values
  values[values < 0] <- 0
  vectors <- eigen$vectors
  factorNames <- factors[,identityColumn]
  realization <- matrix(0, nrow = length(factorNames), ncol=size)
  realization <- apply(realization, 2, rnorm)
  values <- values^(1/2)
  bivariate <- diag(values)
  realization <- bivariate %*% realization
  realization <- vectors %*% realization
  cat(" Realization done \n")
  return(realization)
}

#calculate PL for all factors by type and appluing sensitivities
calculatePL <- function(realization, sensitivity, factors){
  cat(" Calculation PL... \t")
  factorNames <- factors[,identityColumn]
  PL <- matrix( 0, dim(realization)[1], dim(realization)[2], dimnames = list(factorNames, NULL))
  for (i in 1:length(factorNames)){
    factor <- factors[i,]
    if ((factor$Group == "RB" || factor$Group == "EB")&&(factor$Type == "Fut")){
      realization[i,] <-  exp(realization[i,])-1
    } else if ((factor$Group == "EQ" || factor$Group == "FX" || factor$Group == "CM")&&
               (factor$Type == "Spot" || factor$Type == "IV")){
      realization[i,] <-  exp(realization[i,])-1
    }
    cat("#")
    PL[i,] <- realization[i,]*sensitivity[i,"Sens1"] + realization[i,]^2 *sensitivity[i,"Sens1"]
  }
  cat(" Calculation PL done \n")
  return(PL)
}

#create matrix with ones by options parameter
generateIdentities <- function(riskfactors, options){
  cat(" Generating Identityes \t")
  factorNames <- riskfactors[,identityColumn] 
  identityMatrix <- matrix(0, nrow = length(options), ncol = length(factorNames),
                   dimnames = list(options, factorNames))
  for (i in 1:length(options)){
    identityMatrix[i, which(riskfactors[,"Group"] == options[i])] <- 1
  }
  cat(" Identity generation done \n")
  return(identityMatrix)
}

#svae to file and close all
saveToFile <- function(directory, fileName, data){
  cat(" Saving data to file... \t")
  fileNameNormalized=normalizePath(file.path(directory,fileName))
  folder = paste("DBQ=",fileNameNormalized)
  constr = paste(driver, folder, "readOnly=FALSE", sep=";")
  channel <- odbcDriverConnect(constr)
  sqlDrop(channel, "Result", errors = FALSE)
  data <- data.frame(data, stringsAsFactors=F)
  sqlSave(channel, data, rownames = F, tablename = "Result")
  cat(" Savings done \n")
  # odbcClose(channel)
  odbcCloseAll()
}