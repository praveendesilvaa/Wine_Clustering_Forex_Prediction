library(openxlsx)    # For reading and writing Excel files
library(neuralnet)   # For building neural networks
library(Metrics)     # For calculating evaluation metrics
library(nnet)        # For building neural networks


# Load exchange rate data from Excel file
exchange_rate_data <- read.xlsx("/Users/praveendesilva/Documents/final 2/ExchangeUSD.xlsx")

# Display the structure and summary of the data
str(exchange_rate_data)
summary(exchange_rate_data)

# Extract the USD/EUR exchange rate column
exchange_rate <- exchange_rate_data[, 3]

# Clean the data by removing missing values
exchange_rate <- na.omit(exchange_rate)

# Define a function to normalize the data
normalize <- function(x) {
  return((x - min(x)) / (max(x) - min(x)))
}

# Define a function to denormalize the data
denormalize <- function(x, original_data) {
  return(x * (max(original_data) - min(original_data)) + min(original_data))
}

# Normalize the exchange rate data
exchange_rate_normalized <- normalize(exchange_rate)

# Split data into training and testing sets
train_data <- exchange_rate_normalized[1:400]       # Take the first 400 data points for training
test_data <- exchange_rate_normalized[401:length(exchange_rate)]   # Take the rest for testing

# Save the normalized training and testing data to an Excel file
write.xlsx(as.data.frame(train_data), file = "train_data_normalized.xlsx", sheetName = "Train Data", rowNames = FALSE)
write.xlsx(as.data.frame(test_data), file = "test_data_normalized.xlsx", sheetName = "Test Data", rowNames = FALSE)

# Initialize variables to store the best models
best_score <- rep(Inf, 12)   # Initialize vector to store the best scores for comparison
best_mlps <- vector("list", 12)   # Initialize list to store the best MLP models

# Experiment with different input delays, hidden layers, nodes, and activation functions
inputDelays <- c(1, 2, 3, 4)   # Different time delays for input data
hiddenLayers <- c(1, 2)         # Number of hidden layers
nodes <- list(c(5), c(10, 10), c(20, 20), c(30, 30), c(40, 40))   # Number of nodes in each hidden layer
activation_functions <- list("logistic", "tanh")   # Activation functions for hidden layers

# Scoring function
calculate_score <- function(rmse, mae, mape, smape) {
  return((rmse + mae + mape + smape) / 4)
}

# Function to normalize input/output matrix
normalize_matrix <- function(matrix) {
  num_cols <- ncol(matrix)
  for (i in 1:num_cols) {
    matrix[,i] <- normalize(matrix[,i])
  }
  return(matrix)
}

# Function to construct input/output matrix for time-series data
# Input:
#   data: Time-series data vector
#   delay: Number of time steps to consider for input data
# Output:
#   List containing input matrix (X) and output vector (y)
construct_input_output_matrix <- function(data, delay) {
  n <- length(data)                                    # Number of data points
  X <- matrix(0, ncol = delay, nrow = n - delay)       # Initialize input matrix
  y <- data[(delay + 1):n]                             # Initialize output vector
  
  # Fill the input matrix and output vector
  for (i in 1:delay) {
    X[, i] <- data[(delay - i + 1):(n - i)]            # Populate input matrix
  }
  return(list("X" = X, "y" = y))                      # Return input matrix and output vector
}

set.seed(123)

# Loop through different combinations of input delays, hidden layers, nodes, and activation functions
for (delay in inputDelays) {   # Iterate over different input delays
  for (hl in hiddenLayers) {   # Iterate over different numbers of hidden layers
    for (node in nodes) {       # Iterate over different configurations of nodes in hidden layers
      for (activation in activation_functions) {   # Iterate over different activation functions
        
        # Construct input/output matrix for training data
        input_output_train <- construct_input_output_matrix(train_data, delay)
        X_train <- input_output_train$X   # Input matrix for training data
        y_train <- input_output_train$y   # Output vector for training data
        
        # Construct input/output matrix for testing data
        input_output_test <- construct_input_output_matrix(test_data, delay)
        X_test <- input_output_test$X     # Input matrix for testing data
        y_test <- input_output_test$y     # Output vector for testing data
        
        # # Print the X_test matrix
        # cat("X_test matrix:\n")
        # print(X_test)
        
        # Define the MLP model
        mlp_model <- neuralnet(y_train ~ ., data = as.data.frame(cbind(y_train, X_train)),
                               hidden = node, linear.output = TRUE, act.fct = activation)
        
        # Make predictions for training and testing data
        train_predicted <- predict(mlp_model, X_train)
        test_predicted <- predict(mlp_model, X_test)
        
        # Calculate statistical indices for training and testing data
        rmseTrainData <- rmse(train_predicted, y_train)
        maeTrainData <- mae(train_predicted, y_train)
        mapeTrainData <- mape(train_predicted, y_train)
        smapeTrainData <- smape(train_predicted, y_train)
        
        rmseTest <- rmse(test_predicted, y_test)
        maeTest <- mae(test_predicted, y_test)
        mapeTest <- mape(test_predicted, y_test)
        smapeTest <- smape(test_predicted, y_test)
        
        # Calculate the combined score
        score <- calculate_score(rmseTest, maeTest, mapeTest, smapeTest)
        
        # Check if the current model has a lower score than the worst model in the list of best models
        if (score < max(best_score)) {
          # Find the index of the worst model among the 12 best models
          worst_index <- which.max(best_score)
          # Update the list of best models
          best_score[worst_index] <- score
          best_mlps[[worst_index]] <- list(mlp_model, rmseTest, maeTest, mapeTest, smapeTest)
        }
      }
    }
  }
}
# Print the X_test matrix
cat("X_test matrix:\n")
print(X_test)
print(y_test)

# Print Testing RMSE values for the best 12 models
cat("\nTesting RMSE values for the best 12 models:\n")
for (i in 1:12) {
  if (!is.null(best_mlps[[i]])) {
    cat("Model", i, "- Testing RMSE:", best_mlps[[i]][[2]], "MAE:", best_mlps[[i]][[3]], "MAPE:", best_mlps[[i]][[4]], "sMAPE:", best_mlps[[i]][[5]], "\n")
  }
}

# Plot the best 12 MLP models using the best RMSE values
plot.new()
par(mfrow = c(4, 3))
for (i in 1:12) {
  if (!is.null(best_mlps[[i]])) {
    model <- best_mlps[[i]][[1]]
    plot(model, rep = "best")
    title(main = paste("Model", i, "- Testing RMSE:", best_mlps[[i]][[2]], "MAE:", best_mlps[[i]][[3]], "MAPE:", best_mlps[[i]][[4]], "sMAPE:", best_mlps[[i]][[5]]))
  }
}

plot.new()

# Extract the best MLP model and its testing results
best_model_index <- which.min(sapply(best_mlps, function(x) x[[2]]))   # Find the index of the best model
best_model <- best_mlps[[best_model_index]][[1]]   # Extract the best MLP model

# Make predictions for testing data
test_output <- predict(best_model, as.data.frame(X_test))
test_output <- denormalize(test_output, exchange_rate)
test_actual <- denormalize(y_test, exchange_rate)

# Save the denormalized testing data and predictions to an Excel file
write.xlsx(as.data.frame(test_actual), file = "test_actual_denormalized.xlsx", sheetName = "Test Actual", rowNames = FALSE)
write.xlsx(as.data.frame(test_output), file = "test_output_denormalized.xlsx", sheetName = "Test Output", rowNames = FALSE)

# Calculate RMSE
rmseTest <- rmse(test_output, test_actual)

# Print RMSE
cat("Best Model Testing RMSE:", rmseTest, "\n")

# Plotting the prediction output vs. desired output
plot(test_actual, test_output, 
     main = "Prediction Output vs. Desired Output",
     xlab = "Desired Output", ylab = "Prediction Output",
     col = "blue", pch = 19)
abline(0, 1, col = "red")

# Plotting the prediction output vs. desired output (Simple Line Chart)
plot(test_actual, type = "l", col = "green", lwd = 2,
     main = "Prediction Output vs. Desired Output (Line Chart)",
     xlab = "Desired Output", ylab = "Prediction Output")
lines(test_output, col = "blue", lwd = 2)
abline(0, 1, col = "red")

# Statistical indices
cat("\n\nStatistical Indices for the Best Model:\n")
cat("Testing RMSE:", rmseTest, "\n")
cat("Testing MAE:", mae(test_output, test_actual), "\n")
cat("Testing MAPE:", mape(test_output, test_actual), "\n")
cat("Testing sMAPE:", smape(test_output, test_actual), "\n")

if (!is.null(best_model)) {
  # Plot the neural network architecture
  plot(best_model)   # Plot the neural network
} else {
  cat("No best model found.")
}

# Extract the best MLP models and their testing results
best_model_index_1 <- which.min(sapply(best_mlps, function(x) x[[2]]))
best_model_1 <- best_mlps[[best_model_index_1]][[1]]

best_model_index_2 <- which.min(sapply(best_mlps[-best_model_index_1], function(x) x[[2]]))
best_model_2 <- best_mlps[[best_model_index_2]][[1]]

# Compare the model complexity
cat("Model 1 Structure:\n")
print(best_model_1)
cat("\nModel 2 Structure:\n")
print(best_model_2)

# Compare the performance
cat("\nPerformance Comparison:\n")
cat("Model 1 - Testing RMSE:", best_mlps[[best_model_index_1]][[2]], "MAE:", best_mlps[[best_model_index_1]][[3]], "MAPE:", best_mlps[[best_model_index_1]][[4]], "sMAPE:", best_mlps[[best_model_index_1]][[5]], "\n")
cat("Model 2 - Testing RMSE:", best_mlps[[best_model_index_2]][[2]], "MAE:", best_mlps[[best_model_index_2]][[3]], "MAPE:", best_mlps[[best_model_index_2]][[4]], "sMAPE:", best_mlps[[best_model_index_2]][[5]], "\n")

