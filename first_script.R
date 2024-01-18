
library(tidyverse)

config <- config::get()
username <- Sys.getenv("USERPROFILE")

#Attendance emails

attendance_data <-   list.files(path = paste0(username,
                                              config[["attendance_loc"]])) %>%
  str_subset('.xlsx') %>%
  paste(username,
        config[["attendance_loc"]],
        .,
        sep = "\\") %>%
  readxl::read_excel()


if (dim(attendance_data)[2] > 1) {

  print(" Make sure you copy-paste the emails of the people attended and completely delete the 'divisions' column")

} else {


  if (!dir.exists(paste(username,
                        config[["staff_counts_loc"]],
                        sep = "\\ll"))) {

    staff <- list.files() %>%
      str_subset('.xlsx')

  } else {

    staff <- list.files(path = config[["staff_counts_loc"]]) %>%
      dplyr::last() %>%
      paste(username,
            config[["staff_counts_loc"]],
            .,
            sep = "\\")

  }

}



#Staff Counts

  staff <- list.files() %>%
    str_subset('.xlsx') %>%
    readxl::read_excel(sheet = 2) %>%
    janitor::clean_names() %>%
    select(primary_email_address,
           division) %>%
    dplyr::rename(Email = primary_email_address) %>%
    mutate(Email = tolower(Email))



# Minutes

outcome <- merge(emails, staff, all.x = TRUE) %>%
  arrange(division)


writexl::write_xlsx(outcome, paste0(attendance_data_str,
                                   "//",
                                   tt))
