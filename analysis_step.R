# 加载必要的库
library(ggplot2)
library(car)
library(readxl)
library(reshape2)
library(broom)
library(patchwork)
library(lmtest)
library(viridis)
library(broom)
library(dplyr)
library(cowplot)
library(tidyr)
library(kableExtra)


# 加载数据
file_path <- "APH205 Final Report Database v2.xlsx"  # 确保文件路径和名称正确
data <- read_excel(file_path, sheet = "Sheet1")  # 根据你的数据所在的工作表调整
#查看数据情况
str(data)
summary(data)

# 删除含有过多缺失值的行
data <- data[rowSums(is.na(data)) <= 5, ]
# 填补缺失值
categorical_vars <- sapply(data, is.factor)
data[, categorical_vars] <- lapply(data[, categorical_vars], function(x) ifelse(is.na(x), names(which.max(table(x))), x))
# 检查异常值
par(mfrow=c(1, 3))
boxplot(data$ISI_score_F4, main = "ISI Score at Follow-up 4", horizontal = TRUE)
boxplot(data$PHQ9_score_F4, main = "PHQ9 Score at Follow-up 4", horizontal = TRUE)
boxplot(data$GAD7_score_F4, main = "GAD7 Score at Follow-up 4", horizontal = TRUE)
# 定义一个函数用于缩限异常值
limit_outliers <- function(x, lower_quantile = 0.01, upper_quantile = 0.99) {
  quantiles <- quantile(x, probs = c(lower_quantile, upper_quantile), na.rm = TRUE)
  x[x < quantiles[1]] <- quantiles[1]
  x[x > quantiles[2]] <- quantiles[2]
  return(x)
}
# 对PHQ-9, GAD-7, 和ISI得分进行异常值处理
data$PHQ9_score_F4 <- limit_outliers(data$PHQ9_score_F4)
data$GAD7_score_F4 <- limit_outliers(data$GAD7_score_F4)
data$ISI_score_F4 <- limit_outliers(data$ISI_score_F4)
# 归一化处理函数
normalize <- function(x) {
  return((x - min(x, na.rm = TRUE)) / (max(x, na.rm = TRUE) - min(x, na.rm = TRUE)))
}
# 对归一化后的数据进行异常值处理的变量进行归一化
data$PHQ9_score_F4 <- normalize(data$PHQ9_score_F4)
data$GAD7_score_F4 <- normalize(data$GAD7_score_F4)
data$ISI_score_F4 <- normalize(data$ISI_score_F4)
# 检查处理后的数据
summary(data$PHQ9_score_F4)
summary(data$GAD7_score_F4)
summary(data$ISI_score_F4)
# 计算BMI
data$Height_M <- data$Height_BL / 100  # 转换身高从厘米到米
data$BMI <- data$Weight_BL / (data$Height_M^2)
data$Sex <- ifelse(data$Sex == 1, 0, data$Sex)
data$Sex <- ifelse(data$Sex == 2, 1, data$Sex)
data$Sex_name <- factor(data$Sex, levels = c(0, 1), labels = c("Male", "Female"))
data$Age_Group <- cut(data$Age_BL, breaks=c(12,14,16), labels=c("12-14", "15-16"), include.lowest=TRUE)

# 准备变量列表，假设数据集已经包含了这些变量
variables <- c("Academic_Year","Sibling_BL","Morbidity_S_BL","Morbidity_P_BL","Academic_PLS_BL","COVID19_impact_dailylife", "COVID19_impact_studypractice", 
               "COVID19_impact_familyincome", "COVID19_impact_familyhealth", 
               "COVID19_impact_relationship", "PHQ9_cat_BL", "GAD7_cat_BL", "ISI_cat_BL", "BMI", "Sex", "Age_BL")


# 正态性检验和QQ图
test_normality_and_qqplot <- function(data, variable) {
  shapiro_test <- shapiro.test(data[[variable]])
  print(paste("Shapiro-Wilk Test for", variable, ": W =", shapiro_test$statistic, ", p-value =", shapiro_test$p.value))
  qqPlot(data[[variable]], main=paste("QQ Plot for", variable))
}
lapply(variables, function(var) test_normality_and_qqplot(data, var))


# 计算Spearman相关性矩阵
# 完整分析，筛选高度相关的变量
variables <- c("Academic_Year","Sibling_BL","Morbidity_S_BL","Morbidity_P_BL","Academic_PLS_BL","COVID19_impact_dailylife", "COVID19_impact_studypractice", 
               "COVID19_impact_familyincome", "COVID19_impact_familyhealth", 
               "COVID19_impact_relationship", "PHQ9_cat_BL", "GAD7_cat_BL", "ISI_cat_BL", "BMI", "Sex", "Age_BL")
analysis_data <- data[, variables]
spearman_correlation <- cor(analysis_data, method = "spearman", use = "complete.obs")
library(reshape2)
melted_corr <- melt(spearman_correlation)
# 使用ggplot2绘制热图
library(ggplot2)
ggplot(melted_corr, aes(x = Var2, y = Var1, fill = value)) +
  geom_tile(color = "white") +
  scale_fill_binned(option = "C", direction = -1) +  # 使用Viridis颜色方案
  # geom_text(aes(label = sprintf("%.2f", value)), color = "white", size = 3) +  # 添加文本标签
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1),  # 调整文字角度以防重叠
        axis.text.y = element_text(angle = 0, vjust = 0.5, hjust = 1)) +
  labs(x = "", y = "", title = "Spearman Correlation Matrix") +
  coord_fixed()  # 保持x和y轴的一致比例
#选择有强相关性的变量重新分析
variables <- c("Academic_PLS_BL","COVID19_impact_dailylife", "COVID19_impact_studypractice", 
               "COVID19_impact_familyincome", "COVID19_impact_familyhealth", 
               "COVID19_impact_relationship", "PHQ9_cat_BL", "GAD7_cat_BL", "ISI_cat_BL", "Sex")
analysis_data <- data[, variables]
spearman_correlation <- cor(analysis_data, method = "spearman", use = "complete.obs")
library(reshape2)
melted_corr <- melt(spearman_correlation)
# 使用ggplot2绘制热图
library(ggplot2)
ggplot(melted_corr, aes(x = Var2, y = Var1, fill = value)) +
  geom_tile(color = "white") +
  scale_fill_viridis(option = "G", direction = -1) +  # 使用Viridis颜色方案
  geom_text(aes(label = sprintf("%.2f", value)), color = "white", size = 3) +  # 添加文本标签
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1),  # 调整文字角度以防重叠
        axis.text.y = element_text(angle = 0, vjust = 0.5, hjust = 1)) +
  labs(x = "", y = "", title = "Selected Spearman Correlation Matrix") +
  coord_fixed()  # 保持x和y轴的一致比例

# GLM分析
glm_PHQ9 <- glm(PHQ9_score_F4 ~ COVID19_impact_dailylife + COVID19_impact_studypractice +
                  COVID19_impact_familyincome + COVID19_impact_familyhealth +
                  COVID19_impact_relationship + Sex + Age_BL + BMI+ Academic_Year + Sibling_BL + Academic_PLS_BL + Morbidity_S_BL+ Morbidity_P_BL,
                family = gaussian(link = "identity"), data = data)
summary(glm_PHQ9)

glm_GAD7 <- glm(GAD7_score_F4 ~ COVID19_impact_dailylife + COVID19_impact_studypractice +
                  COVID19_impact_familyincome + COVID19_impact_familyhealth +
                  COVID19_impact_relationship + Sex + Age_BL + BMI+ Academic_Year + Sibling_BL + Academic_PLS_BL + Morbidity_S_BL+ Morbidity_P_BL,
                family = gaussian(link = "identity"), data = data)
summary(glm_GAD7)

glm_ISI <- glm(ISI_score_F4 ~ COVID19_impact_dailylife + COVID19_impact_studypractice +
                 COVID19_impact_familyincome + COVID19_impact_familyhealth +
                 COVID19_impact_relationship + Sex + Age_BL + BMI+ Academic_Year + Sibling_BL + Academic_PLS_BL + Morbidity_S_BL+ Morbidity_P_BL,
               family = gaussian(link = "identity"), data = data)
summary(glm_ISI)
# 提取模型结果并整理
phq9_results <- tidy(glm_PHQ9, conf.int = TRUE)
gad7_results <- tidy(glm_GAD7, conf.int = TRUE)
isi_results <- tidy(glm_ISI, conf.int = TRUE)
# 添加模型标识变量
phq9_results$model <- "PHQ-9"
gad7_results$model <- "GAD-7"
isi_results$model <- "ISI"
# 合并数据
combined_results <- rbind(phq9_results, gad7_results, isi_results)
# 绘制综合图
ggplot(combined_results, aes(x = term, y = estimate, color = model)) +
  geom_point(position = position_dodge(width = 0.5)) +
  geom_errorbar(aes(ymin = conf.low, ymax = conf.high), width = 0.2, position = position_dodge(width = 0.5)) +
  coord_flip() +  # 翻转坐标轴使变量名垂直显示
  labs(title = "Model Coefficients Comparison", x = "Variables", y = "Estimate") +
  facet_wrap(~ model, scales = "free_y", ncol = 1) +  # 使用free_y参数并设置为单列显示
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 90, hjust = 1),  # 调整文字角度和对齐方式
        strip.background = element_blank(),  # 移除分面标签的背景
        strip.text.x = element_text(size = 12, face = "bold"))  # 调整分面标签的文字大小和样式
# GLM表格
phq9_results <- tidy(glm_PHQ9) %>%
  mutate(Model = "PHQ-9")
gad7_results <- tidy(glm_GAD7) %>%
  mutate(Model = "GAD-7")
isi_results <- tidy(glm_ISI) %>%
  mutate(Model = "ISI")
# 合并结果
combined_results <- bind_rows(phq9_results, gad7_results, isi_results)
# 选择需要显示的列
combined_results <- combined_results %>%
  select(Model, term, estimate, std.error, statistic, p.value)
# 格式化表格
combined_results %>%
  kable(format = "html", table.attr = "style='width:100%;'") %>%
  kable_styling(bootstrap_options = c("striped", "hover", "condensed", "responsive")) %>%
  add_header_above(c(" " = 1, "Model Results" = 5)) 
# (只显示PHQ-9)选择需要显示的列
combined_results <- combined_results %>%
  filter(Model == "PHQ-9") %>%  # 只选择PHQ-9模型的结果
  select(Model, term, estimate, std.error, statistic, p.value)
# 格式化表格
combined_results %>%
  kable(format = "html", table.attr = "style='width:100%;'") %>%
  kable_styling(bootstrap_options = c("striped", "hover", "condensed", "responsive")) %>%
  add_header_above(c("PHQ-9 Model Results" = 6))  # 只为PHQ-9模型添加一个表头



#分层分析
# 创建交互项模型
glm_PHQ9_interaction <- glm(PHQ9_score_F4 ~ COVID19_impact_dailylife * Sex + 
                              COVID19_impact_studypractice * Sex + 
                              COVID19_impact_familyincome * Sex + 
                              COVID19_impact_familyhealth * Sex + 
                              COVID19_impact_relationship * Sex +
                              COVID19_impact_dailylife * Age_Group + 
                              COVID19_impact_studypractice * Age_Group + 
                              COVID19_impact_familyincome * Age_Group + 
                              COVID19_impact_familyhealth * Age_Group + 
                              COVID19_impact_relationship * Age_Group + 
                              Sex + Age_Group + BMI,
                            family = gaussian(link = "identity"), data = data)

summary(glm_PHQ9_interaction)
# 分别对性别和年龄组进行分层回归分析
sex_groups <- split(data, data$Sex)
age_groups <- split(data, data$Age_Group)
# 定义函数进行回归分析
run_glm <- function(data_subset) {
  glm(PHQ9_score_F4 ~ COVID19_impact_dailylife + COVID19_impact_studypractice +
        COVID19_impact_familyincome + COVID19_impact_familyhealth +
        COVID19_impact_relationship + BMI,
      family = gaussian(link = "identity"), data = data_subset)
}
# 对每个性别组进行回归分析
glm_results_sex <- lapply(sex_groups, run_glm)
# 对每个年龄组进行回归分析
glm_results_age <- lapply(age_groups, run_glm)
# 输出每个组的模型摘要
lapply(glm_results_sex, summary)
lapply(glm_results_age, summary)
# 提取模型结果并添加组标签
tidy_results_sex <- lapply(names(glm_results_sex), function(group) {
  tidy(glm_results_sex[[group]], conf.int = TRUE) %>% mutate(Group = group)
})
tidy_results_age <- lapply(names(glm_results_age), function(group) {
  tidy(glm_results_age[[group]], conf.int = TRUE) %>% mutate(Group = group)
})
# 合并结果
combined_results <- do.call(rbind, c(tidy_results_sex, tidy_results_age))
combined_results$Group <- recode(combined_results$Group, `0` = "Male", `1` = "Female", `12-14` = "Age 12-14", `15-16` = "Age 15-16")
# 绘制回归系数的置信区间图
ggplot(combined_results, aes(x = term, y = estimate, color = Group)) +
  geom_point(position = position_dodge(width = 0.5)) +
  geom_errorbar(aes(ymin = conf.low, ymax = conf.high), width = 0.2, position = position_dodge(width = 0.5)) +
  coord_flip() +
  theme_minimal() +
  labs(title = "Model Coefficients by Group", x = "Variables", y = "Estimate")


# 残差分析
par(mfrow=c(2, 2))
plot(glm_GAD7)
# Cook's Distance，查看数据点的影响
cooksd <- cooks.distance(glm_GAD7)
plot(cooksd, type="h", main="Cook's Distance")
# 确定哪些点的Cook's Distance值过高
abline(h = 4/(nrow(data)-length(coef(glm_GAD7))), col="red")
# 影响图，查看杠杆值和学生化残差
influencePlot(glm_GAD7)
# QQ图检查残差正态性
qqnorm(residuals(glm_GAD7, type="deviance"))
qqline(residuals(glm_GAD7, type="deviance"))

# 回归结果可视化
# 使用 broom 包提取模型的系数和置信区间
tidy_glm <- tidy(glm_GAD7, conf.int = TRUE)
# 绘制系数及其置信区间
coef_plot <- ggplot(tidy_glm, aes(x = term, y = estimate)) +
  geom_point() +
  geom_errorbar(aes(ymin = conf.low, ymax = conf.high), width = 0.2) +
  coord_flip() +
  labs(x = "Variables", y = "Estimate", title = "GLM Coefficients with Confidence Intervals") +
  theme_minimal()
# 显示系数图表
print(coef_plot)
# 计算预测值
predictions <- predict(glm_GAD7, type = "response", newdata = data)
# 绘制预测值与实际值的散点图
scatter_plot <- ggplot(data, aes(x = predictions, y = PHQ9_score_F4)) +
  geom_point(alpha = 0.5) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(x = "Predicted Values", y = "Actual Values", title = "Predictions vs. Actual Values") +
  theme_minimal()
# 显示预测值与实际值的散点图
print(scatter_plot)

# 准备数据，假设需要展示COVID-19影响因素对PHQ-9得分的分布
violin_data <- data[, c("PHQ9_score_F4","PHQ9_score_BL", "COVID19_impact_dailylife", "COVID19_impact_studypractice", 
                        "COVID19_impact_familyincome", "COVID19_impact_familyhealth", 
                        "COVID19_impact_relationship", "Sex", "Age_BL", "BMI")]
# 创建提琴图，展示COVID-19对日常生活影响和PHQ-9得分（对比F4和BL）的分布
#library(devtools)
#install_github("JanCoUnchained/ggunchained")
library(ggunchained)
data_selected <- data %>%
  select(COVID19_impact_dailylife, PHQ9_score_BL, PHQ9_score_F4) %>%
  pivot_longer(cols = c(PHQ9_score_BL, PHQ9_score_F4), names_to = "Group", values_to = "Score") %>%
  mutate(Group = recode(Group, "PHQ9_score_BL" = "Baseline", "PHQ9_score_F4" = "Follow-up 4"))
# 将COVID19_impact_dailylife转为因子类型
data_selected$COVID19_impact_dailylife <- factor(data_selected$COVID19_impact_dailylife)
# 创建图形
ggplot(data_selected, aes(x = COVID19_impact_dailylife, y = Score, fill = Group))+
  geom_split_violin(trim = T,colour="white", scale = 'width')+
  scale_fill_manual(values = c("#1ba7b3","#dfb424"))

p <- ggplot(data_selected, aes(x = COVID19_impact_dailylife, y = Score, fill = Group)) +
  geom_split_violin(trim = FALSE, color = NA, adjust = 1.5, alpha = 0.7, scale = 'width') +
  guides(fill = guide_legend(title = "Time Point")) +
  stat_summary(fun.data = "mean_sd", position = position_dodge(0.15), geom = "errorbar", width = 0.1) +
  stat_summary(fun = "mean", geom = "point", position = position_dodge(0.15), show.legend = FALSE) +
  stat_compare_means(aes(group = Group), label = "p.signif", method = "t.test", size = 5) +
  scale_fill_manual(values = c("#788FCE", "#E6956F")) +
  labs(x = "COVID-19 Impact on Daily Life", y = "PHQ-9 Score") +
  theme(
    axis.text.x = element_text(angle = 0, hjust = 0.5, vjust = 0.5, color = "black", size = 10, margin = margin(b = 2)),
    axis.text.y = element_text(color = "black", size = 10, margin = margin(r = 1)),
    panel.background = element_rect(fill = NA, color = NA),
    panel.grid.minor = element_line(size = 0.2, color = "#e5e5e5"),
    panel.grid.major = element_line(size = 0.2, color = "#e5e5e5"),
    panel.border = element_rect(fill = NA, color = "black", size = 1, linetype = "solid"),
    legend.key = element_blank(),
    legend.title = element_blank(),
    legend.text = element_text(color = "black", size = 8),
    legend.spacing.x = unit(0.1, 'cm'),
    legend.key.width = unit(0.5, 'cm'),
    legend.key.height = unit(0.5, 'cm'),
    legend.position = "none",
    legend.background = element_blank()
  )
# 打印图形
print(p)

# 创建长格式的数据框
long_data <- data %>%
  select(Sex, Age_Group, PHQ9_score_F4, GAD7_score_F4, ISI_score_F4) %>%
  pivot_longer(cols = c(PHQ9_score_F4, GAD7_score_F4, ISI_score_F4),
               names_to = "Test", values_to = "Score")
# 确保数据框中没有NA值
long_data <- na.omit(long_data)
# 绘制密度图
density_plot <- ggplot(long_data, aes(x = Score, fill = interaction(Sex, Age_Group))) +
  geom_density(alpha = 0.7) +
  facet_wrap(~ Test, scales = "free") +
  theme_minimal() +
  labs(title = "Density Plot of Psychological Scores by Sex and Age Group",
       x = "Score", y = "Density") +
  scale_fill_viridis_d(name = "Sex and Age Group",
                    labels = c("Male.12-14", "Female.12-14", "Male.15-16", "Female.15-16"))
print(density_plot)
# 绘制提琴图
violin_plot <- ggplot(long_data, aes(x = interaction(Sex, Age_Group), y = Score, fill = interaction(Sex, Age_Group))) +
  geom_violin(alpha = 0.7) +
  geom_boxplot(width = 0.1, fill = "white") +
  facet_wrap(~ Test, scales = "free") +
  theme_minimal() +
  labs(title = "Violin Plot of Psychological Scores by Sex and Age Group",
       x = "Sex and Age Group", y = "Score") +
  scale_fill_viridis_d(name = "Sex and Age Group",
                    labels = c("Male.12-14", "Female.12-14", "Male.15-16", "Female.15-16"))
print(violin_plot)
# 组合密度图和提琴图
combined_plot <- plot_grid(density_plot, violin_plot, labels = c("A", "B"), ncol = 1)
# 打印组合图
print(combined_plot)


# 将数据转为长格式
data_long <- data %>%
  pivot_longer(cols = c(PHQ9_cat_BL, PHQ9_cat_F4, GAD7_cat_BL, GAD7_cat_F4, ISI_cat_BL, ISI_cat_F4),
               names_to = "variable",
               values_to = "value")
# 添加心理测试类型和时间点信息
data_long <- data_long %>%
  mutate(score_type = case_when(
    grepl("PHQ9", variable) ~ "PHQ-9",
    grepl("GAD7", variable) ~ "GAD-7",
    grepl("ISI", variable) ~ "ISI"
  ),
  time_point = case_when(
    grepl("_BL", variable) ~ "Baseline",
    grepl("_F4", variable) ~ "Follow-up 4"
  ))
# 绘制直方图
ggplot(data_long, aes(x = value, fill = time_point)) +
  geom_histogram(position = "dodge", bins = 20, alpha = 0.7) +
  facet_wrap(~score_type, scales = "free_y") +
  labs(title = "Comparison of Psychological Test Scores at Baseline and Follow-Up 4",
       x = "Score", y = "Frequency") +
  scale_fill_manual(values = c("Baseline" = "skyblue", "Follow-up 4" = "orange")) +
  theme_minimal()

# 绘制叠加直方图
ggplot(data_long, aes(x = value, fill = time_point)) +
  geom_histogram(position = "identity", alpha = 0.5, bins = 30) +
  facet_wrap(~score_type, scales = "free") +
  labs(title = "Comparison of Psychological Test Catagories at Baseline and Follow-Up 4",
       x = "Score", y = "Frequency") +
  scale_fill_manual(values = c("Baseline" = "skyblue", "Follow-up 4" = "orange")) +
  theme_minimal() +
  theme(
    strip.text = element_text(size = 12),
    plot.title = element_text(size = 16, face = "bold"),
    axis.title = element_text(size = 14),
    axis.text = element_text(size = 12)
  )

# 绘制密度图
data_long1 <- data %>%
  pivot_longer(cols = c(PHQ9_score_BL, PHQ9_score_F4, GAD7_score_BL, GAD7_score_F4, ISI_score_BL, ISI_score_F4),
               names_to = "variable",
               values_to = "value")
# 添加心理测试类型和时间点信息
data_long1 <- data_long1 %>%
  mutate(score_type = case_when(
    grepl("PHQ9", variable) ~ "PHQ-9",
    grepl("GAD7", variable) ~ "GAD-7",
    grepl("ISI", variable) ~ "ISI"
  ),
  time_point = case_when(
    grepl("_BL", variable) ~ "Baseline",
    grepl("_F4", variable) ~ "Follow-up 4"
  ))
ggplot(data_long1, aes(x = value, fill = time_point)) +
  geom_density(alpha = 0.5) +
  facet_wrap(~score_type, scales = "free") +
  labs(title = "Density Plot of Psychological Test Scores at Baseline and Follow-Up 4",
       x = "Score", y = "Density") +
  scale_fill_manual(values = c("Baseline" = "skyblue", "Follow-up 4" = "orange")) +
  theme_minimal() +
  theme(
    strip.text = element_text(size = 12),
    plot.title = element_text(size = 16, face = "bold"),
    axis.title = element_text(size = 14),
    axis.text = element_text(size = 12)
  )



