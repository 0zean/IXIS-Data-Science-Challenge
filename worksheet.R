library(openxlsx)
library(tidyverse)
library(dplyr)

# Import CSVs
adds_to_cart <- read.csv("Data\\DataAnalyst_Ecom_data_addsToCart.csv")
session_counts <- read.csv("Data\\DataAnalyst_Ecom_data_sessionCounts.csv")

# Renaming device category column for appearance
session_counts <- session_counts %>%
    rename(DeviceCategory = dim_deviceCategory)

summary(session_counts)

# Clean session_counts data
# session_counts <- session_counts %>%
#     filter(browser != "error" & browser != "(not set)")

# Convert date column to proper date format
session_counts$dim_date <- as.Date(session_counts$dim_date, format = "%m/%d/%y")

# Extract month from dim_date
session_counts$Month <- format(session_counts$dim_date, "%Y-%m")

# Convert sessions, transactions, and QTY columns to numeric
session_counts$sessions <- as.numeric(session_counts$sessions)
session_counts$transactions <- as.numeric(session_counts$transactions)
session_counts$QTY <- as.numeric(session_counts$QTY)


# Aggregating data by Month and Device Category
agg_data <- session_counts %>%
    group_by(Month, DeviceCategory) %>%
    summarize(Sessions = sum(sessions),
            Transactions = sum(transactions),
            QTY = sum(QTY),
            ECR = sum(transactions) / sum(sessions),
            .groups = "drop")

# Creating the first worksheet
wb <- createWorkbook()
addWorksheet(wb, "Month_Device_Aggregation")
writeData(wb, sheet = "Month_Device_Aggregation", agg_data)


# Calculate Month over Month comparison for each variable
agg_month <- session_counts %>%
    group_by(Month) %>%
    summarize(Sessions = sum(sessions),
            Transactions = sum(transactions),
            QTY = sum(QTY),
            ECR = sum(transactions) / sum(sessions),
            .groups = "drop")

# Create Month variable to left join addsToCart to aggregated sessionCounts
adds_to_cart <- adds_to_cart %>%
    mutate(Month = paste(dim_year, sprintf("%02d", dim_month), sep = "-"))

# Left Join
agg_month <- agg_month %>%
    left_join(select(adds_to_cart, Month, addsToCart), by = "Month")


# Helper function for relative differnece
rel_difference <- function(a, b) {
    ((a - b) / ((a + b) / 2)) * 100
}

comparisons <- agg_month %>%
    mutate(
        rel_MoM_Sessions = rel_difference(Sessions, lag(Sessions)),
        rel_MoM_Transactions = rel_difference(Transactions, lag(Transactions)),
        rel_MoM_QTY = rel_difference(QTY, lag(QTY)),
        rel_MoM_ECR = rel_difference(ECR, lag(ECR)),
        rel_MoM_addsToCart = rel_difference(addsToCart, lag(addsToCart)),
        abs_MoM_Sessions = abs(Sessions - lag(Sessions)),
        abs_MoM_Transactions = abs(Transactions - lag(Transactions)),
        abs_MoM_QTY = abs(QTY - lag(QTY)),
        abs_MoM_ECR = abs(ECR - lag(ECR)),
        abs_MoM_addsToCart = abs(addsToCart - lag(addsToCart))
  )

mm_comparison <- tail(comparisons, n = 2)


# Creating the second worksheet
addWorksheet(wb, "Month_over_Month_Comparison")
writeData(wb, sheet = "Month_over_Month_Comparison", mm_comparison)

# Save the workbook
saveWorkbook(wb, "worksheets.xlsx")
