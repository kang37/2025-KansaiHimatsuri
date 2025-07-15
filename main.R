# 加载包。
library(readxl)
library(cluster)
library(factoextra)
library(dplyr)

# 读入数据。
df <- read_excel("data_raw/抽选案例用.xlsx")

# 必要的数据预处理：选取要用于聚类的列，并转换字符为factor。
df_use <- df %>%
  select(類型:田地) %>%
  mutate(across(where(is.character), as.factor))

# 计算Gower距离。
gower_dist <- daisy(df_use, metric = "gower")

# 层次聚类。
hc <- hclust(gower_dist, method = "ward.D2")

# 聚类图可视化。
png("data_proc/clust_res.png", width = 3000, height = 12000, res = 600)
fviz_dend(hc, k = 4, rect = TRUE, cex = 0.3) + coord_flip()
dev.off()

# 导出结果表格。
# 设置聚类数。
k <- 4
cluster_result <- cutree(hc, k = k)

# 加入原始数据，生成聚类结果表格。
df_clustered <- df %>%
  mutate(Cluster = cluster_result)

# 查看前几行结果。
head(df_clustered)

# 导出结果。
write.csv(df_clustered, "data_proc/clust_res.csv", row.names = F)
