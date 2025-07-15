# Preparation ----
library(readxl)
library(cluster)
library(factoextra)
library(dplyr)
library(readr)
library(stringr)
library(FactoMineR)
library(tidyr)

# Direct category ----
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

# 查看前几行结果。
head(df_clustered)

# 导出结果。
write.csv(df_clustered, "data_proc/clust_res.csv", row.names = F)
