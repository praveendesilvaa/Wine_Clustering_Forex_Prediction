# Load required libraries
library(tidyverse)
library(readxl)
library(NbClust)
library(cluster)
library(fpc)
library(factoextra)
library(writexl)


# Load the dataset
wine_data <- read_excel("/Users/praveendesilva/Documents/20221895_W1985643/20221895_W1985643/Whitewine_v6.xlsx")

boxplot(wine_data, main = "Boxplot of Wine Dataset Before Outlier Removal")

# Check for missing values
sum(is.na(wine_data))

# Preprocessing
# Select only the first 11 attributes
wine_data <- wine_data[,1:11]

# Outlier detection and removal using Z-score
z_scores <- scale(wine_data)

# Set threshold for outlier detection
threshold <- 3

# Remove outliers
wine_data_no_outliers <- wine_data[apply(abs(z_scores) < threshold, 1, all),]

# Create boxplot after removing outliers
boxplot(wine_data_no_outliers, main = "Boxplot of Wine Dataset After Outlier Removal")

# Justification for outlier removal
cat("Justification for Outlier Removal: Outliers can significantly affect the performance of clustering algorithms, especially k-means. Outliers can disproportionately influence the centroids, leading to inaccurate cluster assignments. By removing outliers, we ensure that the clusters formed are more representative of the majority of the data points.\n")

# Scaling
scaled_data <- scale(wine_data_no_outliers)
# Convert scaled_data to a data frame
scaled_data_df <- as.data.frame(scaled_data)

# Write scaled data to an Excel file
write_xlsx(scaled_data_df, "/Users/praveendesilva/Documents/20221895_W1985643/20221895_W1985643/Scaled_Wine_Data.xlsx")

#boxplot after scaling data
boxplot(scaled_data, main = "Boxplot of Wine Dataset After Scaling")

# Determine the number of cluster centers via four automated tools

# NbClust for original dataset
nb <- NbClust(scaled_data, min.nc = 2, max.nc = 10, method = "kmeans")
best_k_nb <- nb$Best.nc[1]
nb_clusters <- cbind(2:10, nb$Best.n[2:10])
cat("Best k using NbClust Method:", best_k_nb, "\n")

# NbClust for original dataset - Euclidean Distance
nb_euclidean <- NbClust(scaled_data, min.nc = 2, max.nc = 10, method = "kmeans", index = "all", distance = "euclidean")
best_k_nb_euclidean <- nb_euclidean$Best.nc[1]
nb_clusters_euclidean <- cbind(2:10, nb_euclidean$Best.n[2:10])
cat("Best k using NbClust Method with Euclidean Distance:", best_k_nb_euclidean, "\n")

# NbClust for original dataset - Manhattan Distance
nb_manhattan <- NbClust(scaled_data, min.nc = 2, max.nc = 10, method = "kmeans", index = "all", distance = "manhattan")
best_k_nb_manhattan <- nb_manhattan$Best.nc[1]
nb_clusters_manhattan <- cbind(2:10, nb_manhattan$Best.n[2:10])
cat("Best k using NbClust Method with Manhattan Distance:", best_k_nb_manhattan, "\n")


# Plot NbClust Method for original dataset
# plot(nb)

# Elbow Method
fviz_nbclust(scaled_data, kmeans, method = 'wss')
wss <- numeric(10)
for (i in 2:10) {
  kmeans_model <- try(kmeans(scaled_data, centers = i, nstart = 25, iter.max = 1000), silent = TRUE)
  if (!inherits(kmeans_model, "try-error") && !is.na(kmeans_model$tot.withinss)) {
    wss[i] <- kmeans_model$tot.withinss
  } else {
    cat("Warning: K-means did not converge for k =", i, "\n")
    wss[i] <- NA
  }
}
min_wss_index <- which.min(wss)
best_k_elbow <- min_wss_index + 1  # Adjusting to match k value
cat("Best k using the Elbow Method:", best_k_elbow, "\n")




# Gap Statistic
gap_stat <- clusGap(scaled_data, FUN = kmeans, K.max = 20, B = 50, verbose = TRUE, iter.max = 1000)
best_k_gap <- maxSE(gap_stat$Tab[,"gap"], gap_stat$Tab[,"SE.sim"], method = "Tibs2001SEmax")
cat("Best k using Gap Statistic Method:", best_k_gap, "\n")

# Plot Gap Statistic
plot(gap_stat, main = "Gap Statistic Method to Find Optimal k")
fviz_gap_stat(gap_stat)

# Silhouette Method
sil_width <- rep(NA, 10)
sil_clusters <- NULL
for (i in 2:10) {
  set.seed(123)
  kmeans_fit <- kmeans(scaled_data, centers = i, nstart = 25, iter.max = 1000) # Increased iter.max
  if (length(unique(kmeans_fit$cluster)) > 1) {
    cluster_assignments <- kmeans_fit$cluster
    sil_width[i] <- cluster.stats(dist(scaled_data), cluster_assignments)$avg.silwidth
    sil_clusters <- rbind(sil_clusters, c(i, sil_width[i]))
  }
}
best_k_sil <- which.max(sil_width[2:10]) + 1
cat("Best k using Silhouette Method:", best_k_sil, "\n")

# Plot Silhouette Method
fviz_nbclust(scaled_data, kmeans, method = 'silhouette')


# Combine all cluster numbers
cluster_numbers <- list(NbClust = nb_clusters,
                        Elbow_Method = best_k_elbow,
                        Gap_Statistic = best_k_gap,
                        Silhouette_Method = best_k_sil)

# Choose the best number of clusters
best_k <- min(best_k_nb, best_k_elbow, best_k_sil, best_k_gap)

# Perform k-means clustering with the best k
set.seed(123)
kmeans_fit <- kmeans(scaled_data, centers = best_k, nstart = 25)

# K-means output
kmeans_fit
#plot
fviz_cluster(kmeans_fit, geom = "point", data = scaled_data, alpha = 0.8)

# Ratio of BSS (Between Cluster Sums of Squares) over TSS (Total Sum of Squares)
BSS <- kmeans_fit$betweenss
TSS <- kmeans_fit$totss
BSS_TSS_ratio <- BSS / TSS
cat("BSS/TSS Ratio:", BSS_TSS_ratio, "\n")

# Internal evaluation metrics: BSS and WSS
cat("Between Cluster Sums of Squares (BSS):", BSS, "\n")
cat("Within Cluster Sums of Squares (WSS):", kmeans_fit$tot.withinss, "\n")

# Cluster centers
cat("Cluster Centers:\n")
cluster_centers <- kmeans_fit$centers
print(cluster_centers)

# Cluster assignments
cat("Cluster Assignments:\n")
cluster_assignments <- kmeans_fit$cluster
print(cluster_assignments)

# Silhouette plot
sil <- silhouette(kmeans_fit$cluster, dist(scaled_data))

# Plot silhouette
plot(sil, col = 1:best_k, border = NA, main = "Silhouette Plot for Wine Dataset",
     xlab = "Silhouette Width", ylab = "Cluster")

# Average silhouette width score
avg_sil_width <- mean(sil[, 3])
cat("Average Silhouette Width Score:", avg_sil_width, "\n")

# Save best cluster value
best_cluster_value <- best_k

