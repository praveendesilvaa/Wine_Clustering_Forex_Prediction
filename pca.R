library(xlsx)
# e. Apply PCA to the dataset
# Scale the data
scaled_data <- scale(wine_data_no_outliers)

# Perform PCA analysis
pca <- prcomp(scaled_data)

summary(pca)

# Visualize PCA
fviz_pca_ind(pca, 
             geom.ind = "point", 
             pointshape = 21, 
             palette = "jco", 
             addEllipses = TRUE)

# Eigenvalues and eigenvectors
eigenvalues <- pca$sdev^2
eigenvectors <- pca$rotation

cat("Eigenvalues:\n")
print(eigenvalues)
cat("\n")

cat("Eigenvectors:\n")
print(eigenvectors)
cat("\n")

# Cumulative score per principal components (PC)
cumulative_score <- cumsum(pca$sdev^2 / sum(pca$sdev^2) * 100)

# Plot cumulative score
plot(cumulative_score, xlab = "Number of Principal Components", 
     ylab = "Cumulative Score", type = "b")

# Identify the number of principal components with cumulative score > 85%
num_components <- which(cumulative_score > 85)[1]

# Create a new transformed dataset with selected principal components
transformed_data <- as.data.frame(-pca$x[, 1:num_components])

# Write the data to an Excel file using openxlsx package
# Save the transformed data to an Excel file
openxlsx::write.xlsx(transformed_data, file = "transformed_data.xlsx")


# # Save the transformed data to an Excel file
# write.xlsx(transformed_data, file = "transformed_data.xlsx")

# Confirm that the file was created
cat("PCA transformed data saved as transformed_data.xlsx\n")

# Head of the transformed dataset
head(transformed_data)

cat("Number of Principal Components with Cumulative Score > 85%:", num_components, "\n")

# # Create a new transformed dataset with selected principal components
# transformed_data <- as.data.frame(pca$x[, 1:num_components])


# Brief discussion for the choice of specific number of PCs
cat("Eigenvalues:\n")
print(eigenvalues)
cat("\n")

cat("Cumulative Score per Principal Components:\n")
print(cumulative_score)
cat("\n")





# f. Find an appropriate k for new kmeans clustering attempt on PCA-based dataset

# Determine the number of cluster centers via four automated tools on PCA-based dataset
# 1. NBclust


nb_pca <- NbClust(transformed_data, min.nc = 2, max.nc = 10, method = "kmeans")
best_k_nbclust <- nb_pca$Best.nc[1]
nb_clusters_pca <- cbind(2:10, nb_pca$Best.n[2:10])


# Elbow Method
wss <- function(k) {
  kmeans(transformed_data, k, nstart = 10 )$tot.withinss
}
k.values <- 1:8

wss_values <- map_dbl(k.values, wss)
plot(k.values, wss_values,
     type="b", pch = 19, frame = FALSE,
     xlab="Number of clusters K",
     ylab="Total within-clusters sum of squares")

# Elbow Method
fviz_nbclust(transformed_data, kmeans, method = 'wss')
elbow_clusters_pca <- NULL
wss_pca <- numeric(10)
for (i in 2:10) {
  kmeans_model_pca <- try(kmeans(transformed_data, centers = i, nstart = 25, iter.max = 1000), silent = TRUE)
  if (!inherits(kmeans_model_pca, "try-error")) {
    if (!is.null(kmeans_model_pca$convergence) && kmeans_model_pca$convergence == 0) {
      wss_pca[i] <- kmeans_model_pca$tot.withinss
      elbow_clusters_pca <- rbind(elbow_clusters_pca, c(i, wss_pca[i]))
    } else {
      cat("Warning: K-means did not converge for k =", i, "\n")
      wss_pca[i] <- NA
    }
  } else {
    cat("Error: K-means failed for k =", i, "\n")
    wss_pca[i] <- NA
  }
}

min_wss_index_pca <- which.min(wss_pca)
best_k_elbow_pca <- min_wss_index_pca + 1 # Adjusting to match k value
cat("Best k using the Elbow Method for PCA-based Dataset:", best_k_elbow_pca, "\n")

# Extract the values without NA
if (!is.null(elbow_clusters_pca)) {
  elbow_clusters_pca <- elbow_clusters_pca[complete.cases(elbow_clusters_pca),]
  
  
  plot(1:10, wss_pca, type = "b", xlab = "Number of Clusters (k)", ylab = "Total Within Sum of Squares",
       main = "Elbow Method for PCA-based Dataset")
  points(best_k_elbow_pca, wss_pca[best_k_elbow_pca], col = "red", cex = 2, pch = 19)
}




# 3. Gap Statistic
set.seed(123)
gap_stat_pca <- tryCatch(
  clusGap(transformed_data, FUN = kmeans, nstart = 25, K.max = 10, B = 50, iter.max = 1000),
  error = function(e) {
    cat("Error occurred during Gap Statistic calculation:", conditionMessage(e), "\n")
    NULL
  }
)

if (!is.null(gap_stat_pca)) {
  gap_clusters_pca <- cbind(2:10, gap_stat_pca$Tab[2:10, "gap"])
  max_gap_index <- which.max(diff(gap_stat_pca$Tab[2:10, "gap"]))
  best_k_gap_pca <- max_gap_index + 1
} else {
  cat("Error: Gap Statistic calculation failed\n")
  gap_clusters_pca <- NULL
  best_k_gap_pca <- NA
}

cat("Gap Statistic suggests", best_k_gap_pca, "as the optimal number of clusters.\n")
fviz_gap_stat(gap_stat_pca)

# 4. Silhouette Method
sil_width_pca <- rep(NA, 10)
for(i in 2:10){
  set.seed(123)
  kmeans_fit_pca_sil <- kmeans(transformed_data, centers = i, nstart = 25)
  if (length(unique(kmeans_fit_pca_sil$cluster)) > 1) {
    cluster_assignments_pca <- kmeans_fit_pca_sil$cluster
    sil_width_pca[i] <- cluster.stats(dist(transformed_data), cluster_assignments_pca)$avg.silwidth
  }
}
best_k_sil_pca <- which.max(sil_width_pca[2:10]) + 1
sil_clusters_pca <- cbind(2:10, sil_width_pca[2:10])
cat("Best k using the Silhouette Method:", best_k_sil_pca, "\n")

fviz_nbclust(transformed_data, kmeans, method = 'silhouette')

# Combine all cluster numbers
cluster_numbers_pca <- list(NbClust = nb_clusters_pca,
                            Elbow_Method = elbow_clusters_pca,
                            Gap_Statistic = gap_clusters_pca,
                            Silhouette_Method = sil_clusters_pca)



# g. Using this new pca-dataset, perform a kmeans analysis using the most favoured k from those “automated” methods.

# Perform k-means clustering with the best k = 2
set.seed(123)
kmeans_fit_pca <- kmeans(transformed_data, centers = 2, nstart = 10)

# K-means output
kmeans_fit
#plot
fviz_cluster(kmeans_fit_pca, geom = "point", data = scaled_data, alpha = 0.8)




# Ratio of BSS (Between Cluster Sums of Squares) over TSS (Total Sum of Squares)
BSS_pca <- kmeans_fit_pca$betweenss
TSS_pca <- kmeans_fit_pca$totss
BSS_TSS_ratio_pca <- BSS_pca / TSS_pca
cat("BSS/TSS Ratio for PCA-based Dataset:", BSS_TSS_ratio_pca, "\n")

# Internal evaluation metrics: BSS and WSS
cat("Between Cluster Sums of Squares (BSS) for PCA-based Dataset:", BSS_pca, "\n")
cat("Within Cluster Sums of Squares (WSS) for PCA-based Dataset:", kmeans_fit_pca$tot.withinss, "\n")

# Cluster centers
cat("Cluster Centers for PCA-based Dataset:\n")
cluster_centers_pca <- kmeans_fit_pca$centers
print(cluster_centers_pca)

# Cluster assignments
cat("Cluster Assignments for PCA-based Dataset:\n")
cluster_assignments_pca <- kmeans_fit_pca$cluster
print(cluster_assignments_pca)

# h. Following this “new” kmeans attempt, provide the silhouette plot which displays how close each point in one cluster is to points in the neighbouring clusters. 

# Silhouette plot
sil_pca <- silhouette(kmeans_fit_pca$cluster, dist(transformed_data))

# Plot silhouette
plot(sil_pca, col = 1:best_k_sil_pca, border = NA, main = "Silhouette Plot for PCA-based Dataset",
     xlab = "Silhouette Width", ylab = "Cluster")

# Average silhouette width score
avg_sil_width_pca <- mean(sil_pca[, 3])
cat("Average Silhouette Width Score for PCA-based Dataset:", avg_sil_width_pca, "\n")


# i. Following the kmeans analysis for this new “pca” dataset, implement and illustrate the Calinski-Harabasz Index.

# Calculate Calinski-Harabasz Index
ch_index <- cluster.stats(dist(transformed_data), kmeans_fit_pca$cluster)$ch

# Print Calinski-Harabasz Index
cat("Calinski-Harabasz Index for PCA-based Dataset:", ch_index, "\n")


# Save the best k values for each method
best_k_values <- c(NbClust = best_k_pca,
                   Elbow_Method = best_k_elbow_pca,
                   Gap_Statistic = best_k_gap_pca,
                   Silhouette_Method = best_k_sil_pca)

# Save the cluster assignments for each method
cluster_assignments <- list(NbClust = cluster_assignments_pca,
                            Elbow_Method = kmeans_fit_pca$cluster,
                            Gap_Statistic = kmeans_fit_pca$cluster,
                            Silhouette_Method = kmeans_fit_pca$cluster)

# Display best k values for each method
cat("\nBest k values for each method:\n")
print(best_k_values)

# Display cluster assignments for each method
cat("\nCluster assignments for each method:\n")
print(cluster_assignments)

