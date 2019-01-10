library(e1071)
library(xlsx)
library(rstudioapi)
library(plyr)

current_path <- getActiveDocumentContext()$path 
setwd(dirname(current_path))
origin_file <- 'results/HVAC24hS16-11-2016--0.xlsx'

# Modelo SVM - possivel adicionar parametros (...)
runSVM <- function(file_name, ...){
  
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
  m <- svm(trainingInput,trainingOutput, ...) # ... para adicionar novos parametros
  pred <- predict(m, testData)
  
  return (pred)
}

#Get real value from the original file, according to index (number of inputs in the tested file)
getReal <- function(file_name, index){ 
  trainData <- read.xlsx(file_name, 2, header = T)
  realVal <- trainData[[index,2]]
  return(realVal)
}

#Calculate absolute percent Error for results dataframe
error <- function(table){
  error <- abs((table[ , ] - table[nrow(table), ]) / table[nrow(table), ]) * 100
  return (error)
}

#-----------Run SVM model for each file, using all possible kernels (library e1071)----------------
#----------------------------and changing epsilon value (vector)-----------------------------------

file_names <- sort(list.files(pattern="xlsx$"), decreasing = T)
epsilonList <- c(0.2, 0.5, 0.8, 1)

# Initialize matrix --> Default parameters: kernel = radial & epsilon = 0.1
radial <- as.matrix(sapply(file_names, runSVM))
linear <- as.matrix(sapply(file_names, runSVM, kernel = "linear"))
poli <- as.matrix(sapply(file_names, runSVM, kernel = "polynomial"))
sig <- as.matrix(sapply(file_names, runSVM, kernel = "sigmoid"))

it <- length(epsilonList)
pbar <- create_progress_bar('text') # Check progress of the whole process
pbar$init(it)

# Run Models for each kernel varying epsilon
for (i in 1:it){
  radial <- cbind(radial, sapply(file_names, runSVM, epsilon = epsilonList[i]))
  linear <- cbind(linear, sapply(file_names, runSVM, kernel = "linear", epsilon = epsilonList[i]))
  poli <- cbind(poli, sapply(file_names, runSVM, kernel = "polynomial", epsilon = epsilonList[i]))
  sig <- cbind(sig, sapply(file_names, runSVM, kernel = "sigmoid", epsilon = epsilonList[i]))
  pbar$step()
}
colnames(radial) <- c("radial: 0.1", paste0("radial: ", epsilonList))
colnames(linear) <- c("linear: 0.1", paste0("linear: ", epsilonList))
colnames(poli) <- c("polynomial: 0.1", paste0("polynomial: ", epsilonList))
colnames(sig) <- c("sigmoid: 0.1", paste0("sigmoid: ", epsilonList))

# Combine all previsions and get true values for each test
predValues <- t(cbind(radial, linear, poli, sig))
index <- c(1:length(file_names)) 
realValues <- as.numeric(mapply(getReal, origin_file, index))
predValues <- rbind(predValues, realValues)

# Calculate absolute percent errors and build results dataframe to save
percentError <- as.matrix(MAPE (predValues))
results <- data.frame(predValues, percentError)
results <- results[-nrow(results), ] #Drop last row
colnames(results) <- c(paste0(9:6, "_Inputs"), paste0("Abs_PercentError_", 9:6, "_Inputs"))
MAPE <- rowMeans(results[, 5:8]) # 5:8 index of columns Abs_PercentError
results[, "MAPE"] <- MAPE # Add column MAPE to data.frame

#---------------------------Save results to original file (new sheet)------------------------------

write.xlsx(results, file = origin_file, sheetName = "Ex3_Compare_Results", col.names = T, 
           row.names = T, append = T)
