# Preparation ----
library(readxl)
library(cluster)
library(factoextra)
library(dplyr)
library(readr)
library(stringr)
library(FactoMineR)
library(tidyr)

# Direct cluster ----
# 读入数据。
df <- read_excel("data_raw/抽选案例用.xlsx") %>% 
  rename("神社寺院" = "神社・寺院") %>% 
  rename_with(.cols = c(町中:田地), .fn = ~paste0("其他_", .x)) %>% 
  mutate(類型 = case_when(grepl("\r\n", 類型) ~ "复合式", TRUE ~ 類型))
# 拆分神社寺院列。
df_sub <- df %>% 
  separate(col = 神社寺院, into = paste0("part_", 1:3), sep = "\r\n") %>%
  select(番号, starts_with("part_")) %>% 
  pivot_longer(-番号, names_to = "key", values_to = "place") %>%
  distinct(番号, place) %>% 
  mutate(value = 1) %>% 
  pivot_wider(names_from = place, values_from = value, values_fill = 0) %>% 
  select(-`NA`, -`0`) %>% 
  rename_with(.cols = -番号, .fn = ~paste0("神社寺院_", .x))
# 合并两个数据。
df <- df %>% 
  select(-神社寺院) %>% 
  left_join(df_sub, by = "番号")

# 必要的数据预处理：选取要用于聚类的列，并转换字符为factor。
df_use <- df %>%
  select(類型:神社寺院_森) %>%
  mutate(across(where(is.character), as.factor))

# 计算Gower距离。
gower_dist <- daisy(df_use, metric = "gower")

# 层次聚类。
hc <- hclust(gower_dist, method = "ward.D2")

# 聚类图可视化。
png("data_proc/clust_res_2.png", width = 3000, height = 12000, res = 600)
fviz_dend(hc, k = 4, rect = TRUE, cex = 0.3) + coord_flip()
dev.off()

# 导出结果表格。
# 设置聚类数。
k <- 4
cluster_result <- cutree(hc, k = k)

# 加入原始数据，生成聚类结果表格。
df_clustered <- df %>%
  mutate(Cluster = cluster_result)

# 导出结果。
write.csv(df_clustered, "data_proc/clust_res.csv", row.names = F)

# PCA before cluster ----
# 提取“其他_”和“神社寺院_”开头的列。
df_others <- df %>% select(starts_with("其他_"))
df_shrine <- df %>% select(starts_with("神社寺院_"))

# 对二元变量可以直接做PCA。
pca_others <- PCA(df_others, graph = FALSE)
pca_shrine <- PCA(df_shrine, graph = FALSE)
plot(pca_others)
plot(pca_shrine)

# 提取前两主成分。
df$其他_pc1 <- pca_others$ind$coord[, 1]
df$其他_pc2 <- pca_others$ind$coord[, 2]
df$神社_pc1 <- pca_shrine$ind$coord[, 1]
df$神社_pc2 <- pca_shrine$ind$coord[, 2]

# 合并基本变量作为最终聚类输入。
df_cluster_input <- df %>%
  select(
    類型, 都道府県, 時間, 文化財レベル, 資材の使用量,
    其他_pc1, 其他_pc2, 神社_pc1, 神社_pc2
  ) %>% 
  mutate(across(where(is.character), as.factor))

# 聚类分析。
gower_dist_pca <- daisy(df_cluster_input, metric = "gower")
hc_pca <- hclust(gower_dist_pca, method = "ward.D2")
# 聚类图可视化。
png("data_proc/clust_res_3.png", width = 3000, height = 12000, res = 600)
fviz_dend(hc_pca, k = 4, rect = TRUE, cex = 0.3) + coord_flip()
dev.off()

# 导出结果表格。
# 设置聚类数。
k <- 4
cluster_result_pca <- cutree(hc_pca, k = k)

# 加入原始数据，生成聚类结果表格。
df_clustered_pca <- df %>%
  mutate(Cluster = cluster_result_pca)

# 导出结果。
write.csv(df_clustered_pca, "data_proc/clust_res_3.csv", row.names = F)

# Comparison ----
library(mclust)
adjustedRandIndex(df_clustered$Cluster, df_clustered_pca$Cluster)
table(df_clustered$Cluster, df_clustered_pca$Cluster)

# CA ----
library(FactoMineR)
library(factoextra)
library(showtext)
showtext_auto()

# 假设 df$cluster 为聚类标签，df$文化財レベル 为文化层级
tab <- table(df_clustered$Cluster, df_clustered$類型)

# 对应分析
res.ca <- CA(tab, graph = FALSE)

# 可视化（行=聚类，列=文化等级）
fviz_ca_biplot(res.ca, repel = TRUE)
