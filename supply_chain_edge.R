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

iqr_fg <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.16.23.xlsx",
                           sheet = "Campus FG")

iqr_fg_data_pre <- read_excel("S:/Supply Chain Projects/Linda Liang/Supply Chain Edge/MSTR manual file upload/IQR FG.xlsx")

##############################################################################################################################################################

################## IQR FG Top 5 ##############
strip_leading_zero <- function(date_str){
  sub("^0", "", date_str)
}

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
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_10_data


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
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_25_data


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
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_30_data

# Location 33
iqr_fg_top_5[2, 11] -> iqr_fg_top_5_location_33
iqr_fg_top_5[c(6, 7, 8, 9, 10), 10] -> iqr_fg_top_5_location_33_sku
iqr_fg_top_5[c(6, 7, 8, 9, 10), 11] -> iqr_fg_top_5_location_33_excess_dollar

data.frame(iqr_fg_top_5_location_33, iqr_fg_top_5_location_33_sku, iqr_fg_top_5_location_33_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_33,
                sku = iqr_fg_top_5_location_33_sku,
                excess_in_dollar = iqr_fg_top_5_location_33_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_33_data


# Location 34
iqr_fg_top_5[2, 14] -> iqr_fg_top_5_location_34
iqr_fg_top_5[c(6, 7, 8, 9, 10), 13] -> iqr_fg_top_5_location_34_sku
iqr_fg_top_5[c(6, 7, 8, 9, 10), 14] -> iqr_fg_top_5_location_34_excess_dollar

data.frame(iqr_fg_top_5_location_34, iqr_fg_top_5_location_34_sku, iqr_fg_top_5_location_34_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_34,
                sku = iqr_fg_top_5_location_34_sku,
                excess_in_dollar = iqr_fg_top_5_location_34_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_34_data


# Location 36
iqr_fg_top_5[14, 2] -> iqr_fg_top_5_location_36
iqr_fg_top_5[c(18, 19, 20, 21, 22), 1] -> iqr_fg_top_5_location_36_sku
iqr_fg_top_5[c(18, 19, 20, 21, 22), 2] -> iqr_fg_top_5_location_36_excess_dollar

data.frame(iqr_fg_top_5_location_36, iqr_fg_top_5_location_36_sku, iqr_fg_top_5_location_36_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_36,
                sku = iqr_fg_top_5_location_36_sku,
                excess_in_dollar = iqr_fg_top_5_location_36_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_36_data


# Location 39
iqr_fg_top_5[14, 5] -> iqr_fg_top_5_location_39
iqr_fg_top_5[c(18, 19, 20, 21, 22), 4] -> iqr_fg_top_5_location_39_sku
iqr_fg_top_5[c(18, 19, 20, 21, 22), 5] -> iqr_fg_top_5_location_39_excess_dollar

data.frame(iqr_fg_top_5_location_39, iqr_fg_top_5_location_39_sku, iqr_fg_top_5_location_39_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_39,
                sku = iqr_fg_top_5_location_39_sku,
                excess_in_dollar = iqr_fg_top_5_location_39_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_39_data


# Location 43
iqr_fg_top_5[14, 8] -> iqr_fg_top_5_location_43
iqr_fg_top_5[c(18, 19, 20, 21, 22), 7] -> iqr_fg_top_5_location_43_sku
iqr_fg_top_5[c(18, 19, 20, 21, 22), 8] -> iqr_fg_top_5_location_43_excess_dollar

data.frame(iqr_fg_top_5_location_43, iqr_fg_top_5_location_43_sku, iqr_fg_top_5_location_43_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_43,
                sku = iqr_fg_top_5_location_43_sku,
                excess_in_dollar = iqr_fg_top_5_location_43_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_43_data


# Location 55
iqr_fg_top_5[14, 11] -> iqr_fg_top_5_location_55
iqr_fg_top_5[c(18, 19, 20, 21, 22), 10] -> iqr_fg_top_5_location_55_sku
iqr_fg_top_5[c(18, 19, 20, 21, 22), 11] -> iqr_fg_top_5_location_55_excess_dollar

data.frame(iqr_fg_top_5_location_55, iqr_fg_top_5_location_55_sku, iqr_fg_top_5_location_55_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_55,
                sku = iqr_fg_top_5_location_55_sku,
                excess_in_dollar = iqr_fg_top_5_location_55_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_55_data



# Location 60
iqr_fg_top_5[14, 14] -> iqr_fg_top_5_location_60
iqr_fg_top_5[c(18, 19, 20, 21, 22), 13] -> iqr_fg_top_5_location_60_sku
iqr_fg_top_5[c(18, 19, 20, 21, 22), 14] -> iqr_fg_top_5_location_60_excess_dollar

data.frame(iqr_fg_top_5_location_60, iqr_fg_top_5_location_60_sku, iqr_fg_top_5_location_60_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_60,
                sku = iqr_fg_top_5_location_60_sku,
                excess_in_dollar = iqr_fg_top_5_location_60_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_60_data


# Location 75
iqr_fg_top_5[26, 2] -> iqr_fg_top_5_location_75
iqr_fg_top_5[c(30, 31, 32, 33, 34), 1] -> iqr_fg_top_5_location_75_sku
iqr_fg_top_5[c(30, 31, 32, 33, 34), 2] -> iqr_fg_top_5_location_75_excess_dollar

data.frame(iqr_fg_top_5_location_75, iqr_fg_top_5_location_75_sku, iqr_fg_top_5_location_75_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_75,
                sku = iqr_fg_top_5_location_75_sku,
                excess_in_dollar = iqr_fg_top_5_location_75_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_75_data


# Location 86
iqr_fg_top_5[26, 5] -> iqr_fg_top_5_location_86
iqr_fg_top_5[c(30, 31, 32, 33, 34), 4] -> iqr_fg_top_5_location_86_sku
iqr_fg_top_5[c(30, 31, 32, 33, 34), 5] -> iqr_fg_top_5_location_86_excess_dollar

data.frame(iqr_fg_top_5_location_86, iqr_fg_top_5_location_86_sku, iqr_fg_top_5_location_86_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_86,
                sku = iqr_fg_top_5_location_86_sku,
                excess_in_dollar = iqr_fg_top_5_location_86_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_86_data



# Location 622
iqr_fg_top_5[26, 8] -> iqr_fg_top_5_location_622
iqr_fg_top_5[c(30, 31, 32, 33, 34), 7] -> iqr_fg_top_5_location_622_sku
iqr_fg_top_5[c(30, 31, 32, 33, 34), 8] -> iqr_fg_top_5_location_622_excess_dollar

data.frame(iqr_fg_top_5_location_622, iqr_fg_top_5_location_622_sku, iqr_fg_top_5_location_622_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_622,
                sku = iqr_fg_top_5_location_622_sku,
                excess_in_dollar = iqr_fg_top_5_location_622_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_622_data


# Location 624
iqr_fg_top_5[26, 11] -> iqr_fg_top_5_location_624
iqr_fg_top_5[c(30, 31, 32, 33, 34), 10] -> iqr_fg_top_5_location_624_sku
iqr_fg_top_5[c(30, 31, 32, 33, 34), 11] -> iqr_fg_top_5_location_624_excess_dollar

data.frame(iqr_fg_top_5_location_624, iqr_fg_top_5_location_624_sku, iqr_fg_top_5_location_624_excess_dollar) %>% 
  dplyr::rename(loc = iqr_fg_top_5_location_624,
                sku = iqr_fg_top_5_location_624_sku,
                excess_in_dollar = iqr_fg_top_5_location_624_excess_dollar) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::relocate(date) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>%
  dplyr::mutate(excess_in_dollar = as.double(excess_in_dollar),
                excess_in_dollar = round(excess_in_dollar, 0),
                excess_in_dollar = dollar(excess_in_dollar)) %>% 
  dplyr::filter(sku != "Grand Total") -> iqr_fg_top_5_location_624_data



rbind(iqr_fg_top_5_location_10_data,
      iqr_fg_top_5_location_25_data,
      iqr_fg_top_5_location_30_data,
      iqr_fg_top_5_location_33_data,
      iqr_fg_top_5_location_34_data,
      iqr_fg_top_5_location_10_data,
      iqr_fg_top_5_location_36_data,
      iqr_fg_top_5_location_39_data,
      iqr_fg_top_5_location_43_data,
      iqr_fg_top_5_location_55_data,
      iqr_fg_top_5_location_60_data,
      iqr_fg_top_5_location_75_data,
      iqr_fg_top_5_location_86_data,
      iqr_fg_top_5_location_622_data,
      iqr_fg_top_5_location_624_data) %>% 
  dplyr::rename("Date" = date,
                "Loc" = loc,
                "SKU" = sku,
                "Excess in Dollar" = excess_in_dollar) -> iqr_fg_top_5_total


################## IQR FG #################

# IQR FG this week
iqr_fg[-1:-2,] -> iqr_fg_data
colnames(iqr_fg_data) <- iqr_fg_data[1, ]
iqr_fg_data[-1, ] -> iqr_fg_data

iqr_fg_data %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(campus_ref, mfg_ref, campus, item_2, description, mfg_loc, mfg_line, category, macro_platform, net_wt_lbs, planner, planner_name, label,
                weighted_unit_cost, sum_of_ss, hard_hold_in, sum_of_on_hand, on_hand_in, on_hand_max_in, sum_of_adjusted_forward_inv_target,
                sum_of_adjusted_forward_inv_max, inv_health, iqr, iqr_hold, upi, upi_hold) %>% 
  dplyr::mutate(date = Sys.Date()) %>% 
  dplyr::mutate(date = paste0(strip_leading_zero(format(date, format = "%m")),
                              "/", format(date, format = "%d/%Y"))) %>% 
  dplyr::relocate(date) -> iqr_fg_data


# IQR FG Previous Week

# you need to convert the data, and delete the oldest date


# Combine two
rbind(iqr_fg_data_pre, iqr_fg_data) -> iqr_fg_data_this_week




##########################################################################################################################################################
######################################################################## export to .xlsx format ###########################################################
##########################################################################################################################################################
##########################################################################################################################################################

writexl::write_xlsx(iqr_fg_top_5_total, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 24/Supply Chain Edge Micro Automation/Automation/IQR FG Top 5.xlsx")







