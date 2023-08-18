library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)
library(rio)
library(scales)

######################################################################## Input Data ##########################################################################
iqr_fg_top_5 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.16.23.xlsx",
                           sheet = "Top 5 Excess SKU per Campus-")


##############################################################################################################################################################

strip_leading_zero <- function(date_str){
  sub("^0", "", date_str)
}


# IQR FG Top 5
iqr_fg_top_5 %>% 
  data.frame() -> iqr_fg_top_5

# Location 10
iqr_fg_top_5[2, 2] -> iqr_fg_top_5_location_10
iqr_fg_top_5[c(6, 7, 8, 9, 10), 1] -> iqr_fg_top_5_location_10_sku
iqr_fg_top_5[c(6, 7, 8, 9, 10), 2] -> iqr_fg_top_5_location_10_excess_dollar

data.frame(iqr_fg_top_5_location_10, iqr_fg_top_5_location_10_sku, iqr_fg_top_5_location_10_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_10,
                sku = iqr_fg_top_5_location_10_sku,
                excess_in_dollar = iqr_fg_top_5_location_10_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) -> iqr_fg_top_5_location_10_data


# Location 25
iqr_fg_top_5[2, 5] -> iqr_fg_top_5_location_25
iqr_fg_top_5[c(6, 7, 8, 9, 10), 4] -> iqr_fg_top_5_location_25_sku
iqr_fg_top_5[c(6, 7, 8, 9, 10), 5] -> iqr_fg_top_5_location_25_excess_dollar

data.frame(iqr_fg_top_5_location_25, iqr_fg_top_5_location_25_sku, iqr_fg_top_5_location_25_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_25,
                sku = iqr_fg_top_5_location_25_sku,
                excess_in_dollar = iqr_fg_top_5_location_25_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) -> iqr_fg_top_5_location_25_data


# Location 30
iqr_fg_top_5[2, 8] -> iqr_fg_top_5_location_30
iqr_fg_top_5[c(6, 7, 8, 9, 10), 7] -> iqr_fg_top_5_location_30_sku
iqr_fg_top_5[c(6, 7, 8, 9, 10), 8] -> iqr_fg_top_5_location_30_excess_dollar

data.frame(iqr_fg_top_5_location_30, iqr_fg_top_5_location_30_sku, iqr_fg_top_5_location_30_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_30,
                sku = iqr_fg_top_5_location_30_sku,
                excess_in_dollar = iqr_fg_top_5_location_30_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) -> iqr_fg_top_5_location_30_data


