#Next two lines run only the first time:
# take out hash tag to run them ( then put it back in)
#install.packages("tidyverse")
#install.packages("config")


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
  readxl::read_excel() %>%
  mutate(Email = tolower(Email))


if (dim(attendance_data)[2] > 1) {

  print(" Make sure you copy-paste the emails of the people attended and completely delete the 'divisions' column")

} else {


  if (!dir.exists(paste(username,
                        config[["staff_counts_loc"]],
                        sep = "\\"))) {

    staff <- list.files() %>%
      str_subset('.xlsx')

  } else {

    staff <- list.files(path =  paste0(username, config[["staff_counts_loc"]])) %>%
      dplyr::first() %>%
      paste(username,
            config[["staff_counts_loc"]],
            .,
            sep = "\\")

  }

}



#Staff Counts

  staff <- staff %>%
    readxl::read_excel(sheet = 2) %>%
    janitor::clean_names() %>%
    select(primary_email_address,
           division) %>%
    dplyr::rename(Email = primary_email_address) %>%
    mutate(Email = tolower(Email))




# Minutes

outcome <- merge(attendance_data, staff, all.x = TRUE) %>%
  arrange(division)


writexl::write_xlsx(outcome, paste0(username,
                                     config[["attendance_loc"]],
                                    "\\",
                                     "emails_attendance.xlsx"))




