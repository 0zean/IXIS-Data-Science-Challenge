library(openxlsx)
library(ggplot2)

# Import workbook
page_one <- read.xlsx("worksheets.xlsx", sheet = 1)
page_two <- read.xlsx("worksheets.xlsx", sheet = 2)

custom_palette <- c("#AAC3E8", "#1F2C8F", "#eba5a7")

# Function to create plots for each variable
variable_plot <- function(variable, y_title, title) {
    ggplot(page_one, aes(x = Month, y = {{variable}}, fill = DeviceCategory)) +
        geom_bar(position = "dodge", stat = "identity", width = 0.6) +
        labs(x = "Month", y = y_title, title = title) +
        scale_y_continuous(labels = scales::comma) +
        scale_fill_manual(values = custom_palette) +
        theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, hjust = 1),
            plot.title = element_text(size = 18, face = "bold", hjust = 0.5),
            plot.subtitle = element_text(size = 14),
            plot.caption = element_text(size = 10),
            legend.position = "right")
}

# Sessions
sessions_title <- "Sessions by Month and Device Category"
sessions_plot <- variable_plot(Sessions, "Sessions", sessions_title)
ggsave("Plots\\sessions_plot.png", plot = sessions_plot, dpi = 300)

# QTY
qty_title <- "QTY by Month and Device Category"
qty_plot <- variable_plot(QTY, "QTY", qty_title)
ggsave("Plots\\qty_plot.png", plot = qty_plot, dpi = 300)

# Transactions
transactions_title <- "Transactions by Month and Device Category"
transactions_plot <- variable_plot(Transactions,
                                  "Transactions",
                                   transactions_title)
ggsave("Plots\\transactions_plot.png", plot = transactions_plot, dpi = 300)

# ECR
ecr_title <- "ECR by Month and Device Category"
ecr_plot <- variable_plot(ECR, "ECR", ecr_title)
ggsave("Plots\\ecr_plot.png", plot = ecr_plot, dpi = 300)
