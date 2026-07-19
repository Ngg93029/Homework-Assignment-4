```{r initial_env_setup, include=FALSE}
knitr::opts_chunk$set(echo = TRUE)
```



#Data Setup



```{r}
rm(list = ls())

install_load_packages_hw4 <- function(packages) {
  installed <- rownames(installed.packages())
  missing   <- setdiff(packages, installed)
  if (length(missing) > 0) install.packages(missing, dependencies = TRUE)
  suppressPackageStartupMessages(
    invisible(lapply(packages, require, character.only = TRUE))
  )
}
```



#Load data

```{r chunk_load_csv}

library(tidyverse)

glassdoor <- read.csv("C:/Users/14709/Desktop/Advanced Data Analytics/Advanced Data Analytics/Homework 1/Homework 4/glassdoortest1.csv")
glassdoor <- glassdoor %>% rename(id = X)
```

# Topic Creation

I used multiple AI sources (Google gemini, copilot, cursor, and claude) to find
topic creations for the glassdoor data. When searching, I took items that are
big topics within my current organization to make it more relateable.

Benefits:
-   Benefits-Health
-   Benefits-PTO & Vacation
-   Benefits-Sick leave & Parental Leave
-   Benefits-Retirement & 401(k)

Work-Life Balance
-Work-Life Balance-Flexibility
-Work-Life Balance- Remote & Hybrid Work

Career Development
-Career Development-Promotions & Advancement
-Career Development-Training & Learning

Job Security
-Job Security & Stability-Layoffs & Furloughs
-Job Security & Stability-Employment Stability

Operations & Process
-Operations & Process-Politics & Silos

Workplace Logistics
-Workplace Logistics-Commute & Location
-Workplace Logisitcs- Facilities & Working Conditions

Company Direction
-Company Direction-Strategy & Vision
-Company Direction-Mergers & Organizational Change

Culture & Enviroment
-Culture & Enviroment-Team & Coworkers
-Culture & Enviroment-Recognition & Respect
-Culture & Enviroment-Company Culture & Morale

Compensation
-Compensation-Bonuses & Incentives
-Compensation-Overtime & Extra Pay



List approach -  Instead of writing 20 different, messy chunks of code for 
each topic, putting them into a centralized list lets R loop through all of them
automatically in just a few lines.This specific REGEX structure lets us target 
the 20 defined buckets from the homework sheet, following the exact formatting 
methods taught in the class coding tutorial.

```{r}
categories_hw4 <- list(
  "Benefits-Health" = "\\b(health|medic|healthcare|insur|dental|vision)\\b",
  "Benefits-PTO & Vacation" = "\\b(vacation|pto|holiday|time off|paid time off)\\b",
  "Benefits-Sick leave & Parental Leave" = "\\b(sick|parental|maternity|paternity|leave)\\b",
  "Benefits-Retirement & 401(k)" = "\\b(401k|401\\(k\\)|retirement|pension)\\b",
  
  "Work-Life Balance-Flexibility" = "\\b(flexib|flex|hours)\\b",
  "Work-Life Balance- Remote & Hybrid Work" = "\\b(remote|hybrid|home|wfh)\\b",
  
  "Career Development-Promotions & Advancement" = "\\b(promot|advanc|career|growth|climb)\\b",
  "Career Development-Training & Learning" = "\\b(train|learn|develop|tuition|class)\\b",
  
  "Job Security & Stability-Layoffs & Furloughs" = "\\b(layoff|lay-off|furlough|fired|severance)\\b",
  "Job Security & Stability-Employment Stability" = "\\b(stabil|secur|steady|safe)\\b",
  
  "Operations & Process-Politics & Silos" = "\\b(politic|silo|bureaucra|red tape)\\b",
  
  "Workplace Logistics-Commute & Location" = "\\b(commut|location|travel|distanc|office)\\b",
  "Workplace Logistics- Facilities & Working Conditions" = "\\b(facilit|condition|cafeteria|gym|build)\\b",
  
  "Company Direction-Strategy & Vision" = "\\b(strateg|vision|direction|leader|exec)\\b",
  "Company Direction-Mergers & Organizational Change" = "\\b(merger|acquisit|reorg|chang)\\b",
  
  "Culture & Environment-Team & Coworkers" = "\\b(team|coworker|co-worker|peopl|colleagu)\\b",
  "Culture & Environment-Recognition & Respect" = "\\b(recogni|respect|valu|appreciat)\\b",
  "Culture & Environment-Company Culture & Morale" = "\\b(cultur|morale|environment|atmospher)\\b",
  
  "Compensation-Bonuses & Incentives" = "\\b(bonu|incentiv|stock|equity|shares)\\b",
  "Compensation-Overtime & Extra Pay" = "\\b(overtime|extra pay|hourly|pay|salary|compen)\\b"
)

```



Instead of copying and pasting this entire block of code separately for "Pros" 
and "Cons", wrapping it in a custom function lets me run it on any text column 
I want using a single command.

The loop automatically cycles through our list of 20 categories, applying  
filters to the data table one-by-one and saving hundreds of lines of manual 
coding.

```{r}
tag_comments_hw4 <- function(text_column) {
  # Pull directly from your loaded 'glassdoor' dataframe
  result_df <- glassdoor %>% select(id, date, title, comments = !!sym(text_column))
  
  # Convert column to character strings safely
  raw_chars <- coalesce(as.character(result_df$comments), "")
  
  # FIX: Translate non-UTF-8 smart quotes safely to clean ASCII apostrophes/spaces
  utf8_clean <- iconv(raw_chars, from = "latin1", to = "UTF-8", sub = " ")
  
  # REFERENCE: Text formatting workflow taken from [Doc 3] 'lowercase-comments' and 'remove-line-breaks'
  clean_vec <- tolower(gsub("[\r\n]", " ", utf8_clean))
  
  for (cat_name in names(categories_hw4)) {
    pattern <- categories_hw4[[cat_name]]
    # REFERENCE: Employs 'stringr::str_detect' instead of base 'grep' [Doc 3]
    result_df[[cat_name]] <- if_else(str_detect(clean_vec, pattern), "Y", "N")
  }
  
  # REFERENCE: Functional logic adapted directly from [Doc 3] 'classify-other-and-stop-timer' chunk
  cat_cols <- names(categories_hw4)
  result_df$Other <- if_else(rowSums(result_df[, cat_cols] == "Y") == 0, "Y", "N")
  
  return(result_df)
}
```

Running this to tag "pros" positive feedback
Running this to tag "cons" negative feedback

```{r}
PROS_REPORT <- tag_comments_hw4("pros")
CONS_REPORT <- tag_comments_hw4("cons")
```


#Export excel file

# Setting up to create my excel file

#The intro was the information I selected for my dataspread sheet in excel.

```{r}
INTRO <- c("Lightera LLC",

         "Data Source: Glassdoor",

         "Data As Of: Q2 2026",

         "Prepared on: 7/18/2026",

         "Prepared by: Natalie Gomez")
```


```{r chunk_init_workbook}

library(openxlsx)

wb_hw4 <- createWorkbook()
```

## Modified to output separate tabs for Pros and Cons 

#I am now building my excel file. I created two seperate tabs for "pros" and 
"cons". I wanted to seperate in different tabs positive and negative feedback.


```{r}
addWorksheet(wb_hw4, "Pros Sheet")
addWorksheet(wb_hw4, "Cons Sheet")

# Style for the intro block
intro_style <- createStyle(fontName = "Arial", fontSize = 11, fontColour = "#1F4E78", textDecoration = "Bold")
addStyle(wb_hw4, "Pros Sheet", style = intro_style, rows = 1:5, cols = 1)
addStyle(wb_hw4, "Cons Sheet", style = intro_style, rows = 1:5, cols = 1)

# Write out the intro block text 
writeData(wb_hw4, "Pros Sheet", INTRO, startRow = 1)
writeData(wb_hw4, "Cons Sheet", INTRO, startRow = 1)

# Header style for the main data columns
header_style_hw4 <- createStyle(fontName = "Arial", fontSize = 11, fontColour = "#FFFFFF", 
                                fgFill = "#1F4E78", textDecoration = "Bold", halign = "Center")

#Shifted start Row to 8 so data sits below the INTRO block cleanly 
writeData(wb_hw4, "Pros Sheet", PROS_REPORT, startRow = 8)
writeData(wb_hw4, "Cons Sheet", CONS_REPORT, startRow = 8)

#Updated rows to 8 to apply header color to the table columns 
addStyle(wb_hw4, "Pros Sheet", style = header_style_hw4, rows = 8, cols = 1:ncol(PROS_REPORT), gridExpand = TRUE)
addStyle(wb_hw4, "Cons Sheet", style = header_style_hw4, rows = 8, cols = 1:ncol(CONS_REPORT), gridExpand = TRUE)
```



# REFERENCE: Interactive workbook formatting to make it easier to identify 

```{r}
addFilter(wb_hw4, "Pros Sheet", row = 8, cols = 1:ncol(PROS_REPORT))
addFilter(wb_hw4, "Cons Sheet", row = 8, cols = 1:ncol(CONS_REPORT))

freezePane(wb_hw4, "Pros Sheet", firstActiveRow = 9, firstActiveCol = 5)
freezePane(wb_hw4, "Cons Sheet", firstActiveRow = 9, firstActiveCol = 5)
```




# Save Workbook

```{r}
saveWorkbook(wb_hw4, "Honeywell_Glassdoor_HW4_Report.xlsx", overwrite = TRUE)
cat("\nSuccess! Honeywell_Glassdoor_HW4_Report.xlsx built accurately.\n")
```









