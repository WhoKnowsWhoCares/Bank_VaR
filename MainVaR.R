setwd("F:/OpUR/Alexander Frantsev")
source("Utility.R")
CovarDirectory="F:/OpUR/Alexander Frantsev"
DataDirectory="C:/Users/reuters/Documents/VaR/extracted"
OutputDirectory = DataDirectory
pattern="^AggVaR.*2014.*xls"
patternForCovar = "^AggVaR.*2016.*xlsm"
filesInCovar=list.files(CovarDirectory)
filesInData=list.files(DataDirectory)
supportFile = grep(patternForCovar,filesInCovar,value=TRUE)[1]
dataFiles=grep(pattern,filesInData,value=TRUE)

#Preparations
File <- "AggVaR_2014.03.20.xls"
Factors <- getFactors(CovarDirectory, "RiskFactors.xls")
CovarMatrix <- getCovar(CovarDirectory, File)
SimulationSize <- 1e5
Quantile <- 0.01
Options <- c("RB","EB","FX","MM")
Output <- matrix(0, ncol = length(Options)+1, nrow = 0,
                 dimnames = list(NULL, c("Data", Options)))

#For single file
# result <- calculateQuantile(Factors, CovarMatrix, Options, CovarDirectory, File, SimulationSize)

# For all files
for (File in dataFiles[1:2]){
  quant <- calculateQuantile(Factors, CovarMatrix, Options, DataDirectory, File, SimulationSize)
  Output <- rbind(Output, c(File, quant))
}
saveToFile(OutputDirectory,"Result.xls",Output)

#Main function that calculate Quantile by Options to File in Directory with Factors and CovarMatrix
calculateQuantile <- function(Factors, CovarMatrix, Options, CovarDirectory, File, SimulationSize){
  SubFactors <- truncFactors(Factors, Options, CovarMatrix)
  SubCovar <- truncCovar(CovarMatrix, SubFactors)
  eigen <- eigen(SubCovar)
  Sensitivities <- getSensitivity(CovarDirectory, File, SubCovar, SubFactors)
  Constants <- getConstants(CovarDirectory, File, Options, SimulationSize)
  Simulation <- simulate(SimulationSize, SubFactors, eigen)
  PL <- calculatePL(Simulation, Sensitivities, SubFactors)
  Identities <- generateIdentities(SubFactors, Options)
  Result <- (Identities %*% PL) + Constants
  SortedResult <- apply(Result, 1, sort)
  quant <- as.double(apply(SortedResult, 2, quantile, probs=c(0.01)))
  return(quant)
}