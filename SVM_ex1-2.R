library(e1071)
library(xlsx)
library(rstudioapi)

current_path <- getActiveDocumentContext()$path 
setwd(dirname(current_path))
origin_file <- 'results/HVAC24hS16-11-2016--0.xlsx'

# Modelo SVM - exemplo sem especificacao do kernel (defeito: radial)
runSVM <- function(file_name){
  
  trainingInput <- read.xlsx(file_name, 1, header=T)
  ncols.trainingInput <- ncol(trainingInput)
  trainingInput <- trainingInput[ , 2 : ncols.trainingInput]
  
  trainingOutput<- read.xlsx(file_name, 2, header=T)
  trainingOutput <- matrix(c(trainingOutput[ , 2]), ncol=1)
  
  testData <- read.xlsx(file_name, 3, header=T)
  ncols.testData <- ncol(testData)
  testData <- testData[, 2 : ncols.testData]
  
  #generate column names for trainingInput and trainingOutput. required for formula construction
  colnames(trainingInput) <- inputColNames <- paste0("x", 1:ncol(trainingInput))
  colnames(trainingOutput) <- outputColNames <- paste0("y", 1:ncol(trainingOutput))
  
  #Column bind the data into one variable
  trainingData <- cbind(trainingInput,trainingOutput)
  
  # estimate model and predict input values
  m <- svm(trainingInput,trainingOutput)
  pred <- predict(m, testData)
  
  return (pred)
}

#Just for visualization (header in excel)
getDates <- function(file_name){ 
  testData <- read.xlsx(file_name, 3, header=T)
  data <- testData[[1]]
  return(data)
}

#Get real value from the original file, according to index (number of inputs in the tested file)
getReal <- function(file_name, index){ 
  trainData <- read.xlsx(file_name, 2, header = T)
  realVal <- trainData[[index,2]]
  return(realVal)
}

#-----------------------------------------EXERCISE 1---------------------------------------------

predEx1 <- runSVM('6_Inputs.xlsx')
trueVal <- getReal(origin_file, 4) # index = 10(total) - 6(inputs)
compare <- rbind(predEx1, trueVal)
compare
((trueVal - predEx1)/trueVal) * 100
#----------------------------------------EXERCISE 2----------------------------------------------

#-----------Run SVM model for each file (different number of inputs & test values)---------------

file_names <- sort(list.files(pattern="xlsx$"), decreasing = T) #Reverse order to get right index in getReal

predValues <- as.matrix(sapply(file_names, runSVM))
datasTestadas <- as.matrix(sapply(file_names, getDates))
index <- c(1:length(file_names)) 
realValues <- as.matrix(mapply(getReal, origin_file, index))
percentError <- abs((realValues - predValues) / realValues) * 100

results <- data.frame(datasTestadas, realValues, predValues, percentError)
colnames(results) <- c("Date", "Real Value", "Predicted value", "Absolute Percent Error")
MAPE <- mean(percentError)
MAPE

#---------------------------Save results (ex.2) to original file--------------------------------

write.xlsx(results, file = origin_file, sheetName = "Results_Errors", col.names = T, 
           row.names = T, append = T)
