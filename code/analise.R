library(tidyverse)
library(xlsx)
library(openxlsx)
library(extrafont)
library(ggthemes)
library(scales)
library(flextable)
library(officer)



#load fonts#

font_import(pattern = 'Playfair', prompt = FALSE)
font_import(pattern = 'OpenSans', prompt = FALSE)

#load tables#

full_data <- read.xlsx('data/full_data.xlsx', 1)

values.objectives <- read.xlsx('data/data.xlsx', 'Values and Objectives') %>% 
  mutate(Result = as.character(Result))
  values.objectives$Result[values.objectives$Result == "T"] <- "X"
  values.objectives$Result[values.objectives$Result == "F"] <- "" 
  
program.list <- read.xlsx('data/data.xlsx', 'Program list')

outcomes.objectives <- read.xlsx('data/data.xlsx', 'Outcomes and Objectives') %>% 
  mutate(Result = as.character(Result))
  outcomes.objectives$Result[outcomes.objectives$Result == "T"] <- "X"
  outcomes.objectives$Result[outcomes.objectives$Result == "F"] <- ""
  
outcomes.programs <- read.xlsx('data/data.xlsx', 'Outcomes and Programs')
  
outcomes.courses <- read.xlsx('data/data.xlsx', 'Courses and Outcomes') %>% 
  replace_na(list( I = "", II = "", III = "", IV = "", V = "", VI = ""))

LO.levels <- read.xlsx('data/data.xlsx', 'Proficiency')

config <- read.xlsx('data/data.xlsx', 'Configuração') %>% 
  mutate(Revaluation = as.logical(Revaluation))

courses <- read.xlsx('data/data.xlsx', 'Courses')

employ <- read.xlsx('data/employer_surveys.xlsx', 1)

students <- read.xlsx('data/student_surveys.xlsx', 1)

AY2223 <- read.xlsx('data/2223.xlsx', 1)

#Set outcomes as factors#

config$Outcome <- as.factor(config$Outcome)

config <- inner_join(config, courses, by='Code')

config <- config[,c(1:3,9,5,4,6:8)]


#define collors#

c.color <- RColorBrewer::brewer.pal(3,"YlGnBu")
t.color <- RColorBrewer::brewer.pal(3,"RdYlGn")

# Define custom levels for Program and LO
program_levels <- c(1, 2, 14, 15, 35, 16, 36, 37, 38, 39, 27, 33)
lo_levels <- c("I", "II", "III", "IV", "V", "VI")

#Values and objectives#

VO_table <- function(a){
  
  #p.name <- program.list$Program[program.list$Program_code == a]
  
  small_border = fp_border(color="gray", width = 1)
  no_border = fp_border(color="gray", width = 0)
  
  values.objectives %>%
    subset(Program == a) %>%
    distinct(Program, Objective, Value, .keep_all = TRUE) %>%
    pivot_wider(names_from = Value, values_from = Result) %>% 
    select(-Program) %>%
    as.data.frame() %>%
    flextable() %>% 
    # theme_alafoli %>% 
    width(j = 1, width=45, unit = "mm") %>%
    # add_header_row(values = p.name, colwidths = 6) %>% 
    # add_header_row(values = "Nova SBE values and Program objectives", colwidths = 6) %>% 
    # font(part = "header", fontname = "Playfair Display", i=c(1:2)) %>% 
    font(part = "body", fontname = "Open Sans") %>% 
    font(part = "header", fontname = "Open Sans", i=1) %>% 
    align(align = "center", part = "body", j=c(2:6)) %>% 
    align(align = "center", part = "header", i=1) %>% 
    bold(part = "body", j=c(2:6)) %>% 
    #italic(part = "header", i=2) %>% 
    fontsize(part = "header", size = 7, i=1) %>% 
    #fontsize(part = "header", size = 14, i=1) %>% 
    fontsize(part = "body", size = 7, j=1) %>% 
    valign(valign = "top", part = "header", i=1) %>% 
    border_inner_h(part="body", border = small_border ) %>% 
    border_inner_h(part="header", border = no_border ) %>%
    border_inner_v(part="body", border = small_border ) %>% 
    valign(valign = "center", part = "body", j=c(2:6))
  
}

#Outcomes and objectives#


OO_table <- function(a){
  
  #p.name <- program.list$Program[program.list$Program_code == a]
  
  outcomes.values <-subset(outcomes.programs ,outcomes.programs$Program == a) %>% 
    select(-Program) %>% 
    unite('Merged',Outcome_Code:Outcome, sep = ". ")
  
  small_border = fp_border(color="gray", width = 1)
  no_border = fp_border(color="gray", width = 0)
  
  outcomes.objectives %>% 
    subset(outcomes.objectives$Program == a) %>% 
    pivot_wider(names_from = Outcome, values_from = Result) %>% 
    select(-Program) %>% 
    as.data.frame() %>% 
    `colnames<-`(c(' ',outcomes.values$Merged)) %>% 
    flextable() %>% 
    # theme_alafoli %>% 
    width(j = 1, width=35, unit = "mm") %>%
    width(j = c(2:7), width=35, unit = "mm") %>%
    #add_header_row(values = p.name, colwidths = 7) %>% 
    #add_header_row(values = "Program Learning Outcomes and Program Objectives", colwidths = 7) %>% 
    #font(part = "header", fontname = "Playfair Display", i=c(1:2)) %>% 
    font(part = "body", fontname = "Open Sans") %>% 
    font(part = "header", fontname = "Open Sans", i=1) %>% 
    align(align = "center", part = "body", j=c(2:7)) %>% 
    align(align = "center", part = "header", i=1) %>% 
    bold(part = "body", j=c(2:7)) %>% 
    #italic(part = "header", i=2) %>% 
    fontsize(part = "header", size = 6, i=1) %>% 
    #fontsize(part = "header", size = 14, i=1) %>% 
    fontsize(part = "body", size = 6, j=1) %>% 
    border_inner_h(part="body", border = small_border ) %>% 
    border_inner_h(part="header", border = no_border ) %>%
    border_inner_v(part="body", border = small_border ) %>% 
    valign(valign = "center", part = "body", j=c(2:7)) %>% 
    valign(valign = "top", part = "header", i=1)

}

#Outcomes per program #


LO_table <- function(a){
  
  #p.name <- program.list$Program[program.list$Program_code == a]
  
  small_border = fp_border(color="gray", width = 1)
  no_border = fp_border(color="gray", width = 0)
  
  subset(outcomes.programs ,outcomes.programs$Program == a) %>% 
    select(-Program) %>%
    rename_with(~ c(' '), 1) %>%
    flextable() %>% 
    # theme_alafoli %>% 
    width(j = 2, width=145, unit = "mm") %>% #era 125
    # add_header_row(values = p.name, colwidths = 2) %>% 
    # add_header_row(values = "Program Learning Outcomes", colwidths = 2) %>% 
    font(part = "header", fontname = "Playfair Display", i=1) %>% 
    font(part = "body", fontname = "Open Sans", j=2) %>% 
    font(part = "body", fontname = "Playfair Display", j=1) %>% 
    bold(part = "body", j=1) %>% 
    # italic(part = "header", i=2) %>% 
    # fontsize(part = "header", size = 16, i=1) %>%
    fontsize(part = "body", size = 16, j=1) %>% 
    fontsize(part = "body", size = 7, j=2) %>% 
    align(align = "center", part = "body", j=1) %>% 
    # align(align = "center", part = "header", i=c(1:2)) %>% 
    border_inner_h(part="body", border = small_border ) %>% 
    border_inner_h(part="header", border = no_border ) %>%
    border_inner_v(part="body", border = small_border ) %>% 
    valign(valign = "center", part = "body", j=1) %>% 
    bg(part="body", i=c(4:6),bg="#dae9ec") %>% 
    footnote(j=1, i=c(4:6), value = as_paragraph("Universal learning Outcomes shared among all Nova SBE programs."), ref_symbols = "1", part = "body") %>% 
    italic(part = "foot", i=1) %>% 
    fontsize(part = "foot", size = 9, i=1) %>% 
    valign(valign = "center", part = "body", j = 1)

}

# Courses and Outcomes #


CO_table <- function(a){
  
  #p.name <- program.list$Program[program.list$Program_code == a]
  
  small_border = fp_border(color="gray", width = 1)
  no_border = fp_border(color="gray", width = 0)
  
  outcomes.courses$Code <- as.character(outcomes.courses$Code)
  
  colormatrix <- subset(outcomes.courses ,outcomes.courses$Program == a) %>% 
    select(-c(1:4))
  
  colormatrix[colormatrix == "Developing"] <- "#F5FBD5"
  colormatrix[colormatrix == "Proficient"] <- "#E2F4EF"
  colormatrix[colormatrix == "Expert"] <- "#CAE2F2"
  colormatrix[colormatrix == ""] <- "#FFFFFF"
  
  
  subset(outcomes.courses ,outcomes.courses$Program == a) %>% 
    select(-Program) %>%
    flextable() %>% 
    # theme_alafoli %>% 
    #add_header_row(values = p.name, colwidths = 9) %>% 
    #add_header_row(values = "Levels of Proficiecy in the Learning Outcomes", colwidths = 9) %>% 
    #font(part = "header", fontname = "Playfair Display", i=c(1:2)) %>% 
    font(part = "body", fontname = "Open Sans") %>% 
    font(part = "header", fontname = "Open Sans", j=c(1:3), i=1) %>%
    font(part = "header", fontname = "Playfair Display", j=c(4:9), i=1) %>% 
    bold(part = "header", j=c(4:9), i=1) %>% 
    #italic(part = "header", i=2) %>% 
    #fontsize(part = "header", size = 14, i=1) %>% 
    #fontsize(part = "header", size = 14, i=1, j=c(4:9)) %>% 
    fontsize(part = "body", size = 6, j=c(4:9)) %>% 
    fontsize(part = "body", size = 7, j=c(1:3)) %>% 
    fontsize(part = "header", size = 7, j=c(1:3), i=1) %>% 
    align(align = "center", part = "body", j=c(2:9)) %>% 
    #align(align = "center", part = "header", i=c(1:2)) %>% 
    align(align = "center", part = "header", i=1, j=c(2:9)) %>% 
    width(j = c(1), width = 100, unit = "mm") %>% 
    width(j = c(2:3), width = 15, unit = "mm") %>% 
    width(j = c(4:9), width = 18, unit = "mm") %>% 
    border_inner_h(part="body", border = small_border ) %>% 
    border_inner_h(part="header", border = no_border ) %>%
    border_inner_v(part="body", border = small_border ) %>%
    bg(bg=t(t(colormatrix)), part="body", j = c(4:9))
  
  
}


#Proficiency table#

Prof_table <- function(a){
  
  #p.name <- program.list$Program[program.list$Program_code == a]
  
  small_border = fp_border(color="gray", width = 1)
  no_border = fp_border(color="gray", width = 0)
  
  subset(LO.levels ,LO.levels$Program == a) %>% 
    select(-Program) %>%
    `names<-`(c(" ", "Developing", "Proficient", "Expert")) %>% 
    flextable() %>% 
    # add_header_row(values = p.name, colwidths = 4) %>% 
    # add_header_row(values = "Courses and Levels of Proficiency in the Learning Outcomes", colwidths = 4) %>% 
    font(part = "header", fontname = "Playfair Display", i=1) %>% 
    font(part = "body", fontname = "Open Sans", j=c(2:4)) %>% 
    font(part = "body", fontname = "Playfair Display", j=1) %>%
    #fontsize(part = "header", size = 14, i=1) %>% 
    fontsize(part = "body", size = 7, j=c(2:4)) %>% 
    fontsize(part = "body", size = 14, j=1) %>% 
    #italic(part = "header", i=2) %>% 
    align(align = "center", part = "body", j=1) %>% 
    align(align = "center", part = "header") %>% 
    border_inner_h(part="body", border = small_border ) %>% 
    border_inner_h(part="header", border = no_border ) %>%
    border_inner_v(part="body", border = small_border ) %>%
    bg(bg="#F5FBD5", part="body", j = 2) %>% 
    bg(bg="#F5FBD5", part="header", i =1, j = 2) %>% 
    bg(bg="#E2F4EF", part="body", j = 3) %>% 
    bg(bg="#E2F4EF", part="header", i =1, j = 3) %>% 
    bg(bg="#CAE2F2", part="body", j = 4) %>% 
    bg(bg="#CAE2F2", part="header", i =1, j = 4) %>% 
    width(j = 1, width=10, unit = "mm") %>% 
    width(j = c(2:4), width=50, unit = "mm")

  
}


#Configuração#

config_table <- function(a){
  
  subset(config ,config$Program == a) %>% 
    select(-c(Program, Academic.Year)) %>%
    mutate(Code = as.character(Code)) %>% 
    data.table::setorder('Outcome') %>%
    mutate(Revaluation = ifelse(Revaluation == FALSE, "", ifelse(Revaluation == TRUE, "X", Revaluation))) %>%
    flextable() %>% 
    font(part = "body", fontname = "Open Sans", j=c(1,3,5)) %>% 
    font(part = "body", fontname = "Playfair Display", j=c(4:5)) %>% 
    fontsize(part = "body", size = 8, j=c(1:3,5:7)) %>% 
    fontsize(part = "header", size = 10) %>% 
    align(align = "center", part = "body", j= c(1,3:7)) %>%
    align(align = "center", part = "header") %>%
    bold(part="body", j = c(4,7)) %>% 
    width(j = c(2,6), "width"=40, unit = "mm") %>%
    width(j = c(1,3:5), "width"=15.5, unit = "mm") %>%
    bg(bg="#F5FBD5", part="body", j = 5, i = ~ Level == "Developing") %>% 
    bg(bg="#E2F4EF", part="body", j = 5, i = ~ Level == "Proficient") %>% 
    bg(bg="#CAE2F2", part="body", j = 5, i = ~ Level == "Expert")
  
  
}


# Gráfico boxplot #

data.course.outcome <- function(a,b,c){
  
  TOT <- subset(full_data ,full_data$Program == a & full_data$Outcome == b & full_data$Revaluation == c) %>% 
    nrow()
  
  BT <- subset(full_data ,full_data$Program == a & full_data$Outcome == b & full_data$Revaluation == c) %>% 
    filter(Score < 12) %>% 
    nrow()
  
  Targ <- subset(full_data ,full_data$Program == a & full_data$Outcome == b & full_data$Revaluation == c) %>% 
    filter(Score >= 12 & Score <= 16) %>% 
    nrow()
  
  AT <- subset(full_data ,full_data$Program == a & full_data$Outcome == b & full_data$Revaluation == c) %>% 
    filter(Score > 16) %>% 
    nrow()
  
  TTAT <- Targ+AT
  
  BT <- percent(BT/TOT, accuracy = .01)
  Targ <- percent(Targ/TOT, accuracy = .01)
  AT <- percent(AT/TOT, accuracy = .01)
  TTAT <- percent(TTAT/TOT, accuracy = .01)
  
  av.score <- subset(full_data ,full_data$Program == a & full_data$Outcome == b & full_data$Revaluation == c) %>%
    select(Score) %>% 
    summarize_at("Score", mean) %>% 
    as.numeric() %>% 
    format(digits = 2, nsmall = 2)
  
  
  
  # p <<- subset(full_data ,full_data$Program == a & full_data$Outcome == b) %>% 
  #   ggplot(aes(x=Score))+
  #   geom_boxplot(aes(x=Score), alpha=.4, fill="#34BBDA", size = 1.5, color ="#34BBDA", 
  #                outlier.shape = 23, outlier.size = 3, outlier.alpha = .4, 
  #                outlier.color = "#34BBDA", outlier.fill= "#34BBDA",
  #                width=0.75)+
  #   scale_x_continuous(limits = c(0,20), breaks = c(0,5,10,15,20))+
  #   scale_y_continuous(limits = c(-1,1))+
  #   geom_vline(aes(xintercept=12),
  #              color="#12326E", linetype=2, size=2)+
  #   geom_vline(aes(xintercept=16),
  #              color="#12326E", linetype=2, size=2)+
  #   geom_text(aes(14,-1),label = "Target", size = 4, color = "#12326E", family = "Open Sans Light")+
  #   geom_text(aes(18,-1),label = "Above Target", size = 4, color = "#12326E", family = "Open Sans Light")+
  #   geom_text(aes(6,-1),label = "Below Target", size = 4, color = "#12326E", family = "Open Sans Light")+
  #   geom_text(aes(14,1),label = Targ, size = 4, color = "#12326E", family = "Open Sans Light")+
  #   geom_text(aes(18,1),label = AT, size = 4, color = "#12326E", family = "Open Sans Light")+
  #   geom_text(aes(6,1),label = BT, size = 4, color = "#12326E", family = "Open Sans Light")+
  #   theme_minimal()+
  #   theme(text=element_text(size=16, family="Open Sans Light"))+
  #   theme(axis.text.y =element_blank(), 
  #         axis.ticks.y=element_blank(), 
  #         panel.grid.minor.y=element_blank(),
  #         panel.grid.major.y=element_blank(),
  #         axis.title.y = element_blank()
  #   )
  
  p <<- subset(full_data, full_data$Program == a & full_data$Outcome == b & full_data$Revaluation == c) %>% 
    ggplot(aes(x = Score)) +
    geom_boxplot(
      aes(x = Score), 
      alpha = .4, 
      fill = "#34BBDA", 
      size = 1.5, 
      color = "#34BBDA", 
      outlier.shape = 23, 
      outlier.size = 3, 
      outlier.alpha = .4, 
      outlier.color = "#34BBDA", 
      outlier.fill = "#34BBDA",
      width = 0.75
    ) +
    scale_x_continuous(limits = c(0, 20), breaks = c(0, 5, 10, 15, 20)) +
    scale_y_continuous(limits = c(-1, 1)) +
    geom_vline(aes(xintercept = 12), color = "#12326E", linetype = 2, size = 2) +
    geom_vline(aes(xintercept = 16), color = "#12326E", linetype = 2, size = 2) +
    annotate("text", x = 14, y = -1, label = "Target", size = 4, color = "#12326E", family = "Open Sans Light") +
    annotate("text", x = 18, y = -1, label = "Above Target", size = 4, color = "#12326E", family = "Open Sans Light") +
    annotate("text", x = 6, y = -1, label = "Below Target", size = 4, color = "#12326E", family = "Open Sans Light") +
    annotate("text", x = 14, y = 1, label = Targ, size = 4, color = "#12326E", family = "Open Sans Light") +
    annotate("text", x = 18, y = 1, label = AT, size = 4, color = "#12326E", family = "Open Sans Light") +
    annotate("text", x = 6, y = 1, label = BT, size = 4, color = "#12326E", family = "Open Sans Light") +
    theme_minimal() +
    theme(
      text = element_text(size = 16, family = "Open Sans Light"),
      axis.text.y = element_blank(), 
      axis.ticks.y = element_blank(), 
      panel.grid.minor.y = element_blank(),
      panel.grid.major.y = element_blank(),
      axis.title.y = element_blank()
    )
  
  
  data.frame('KPIs' = c("# of assessments", "Mean score", "Below Target", "Target", 
                        "Above Target", "Total on Target and Above Target"),
             ' ' = c(TOT, av.score, BT, Targ, AT, TTAT))  %>% 
    rename_with(~ c(' '), 2) %>% 
    flextable() %>% 
    font(part = "body", fontname = "Open Sans") %>% 
    font(part = "header", fontname = "Playfair Display") %>% 
    width(j = 1, "width"=75, unit = "mm") %>% 
    width(j = 2, "width"=20, unit = "mm") %>% 
    bold(part="body", i=6) %>% 
    bg(bg="#FECFB9", part="body", i=3) %>% 
    bg(bg="#FFFFAB", part="body", i=4) %>% 
    bg(bg="#E4F3D9", part="body", i=5) %>% 
    footnote(i = 1, j = 1, 
             value = as_paragraph("Total number of assessments used to evaluate the current Learning Outcome."), 
             part="body", ref_symbols = "1") %>% 
    footnote(i = 2, j = 1, 
             value = as_paragraph("Value on a scale of 0 to 20."), 
             part="body", ref_symbols = "2") %>% 
    fontsize(part = "footer", size = 8) %>% 
    font(part = "footer", fontname = "Open Sans")
  
  
  
  
  
  
  OR_table <<-  data.frame('KPIs' = c("# of assessments", "Mean score", "Below Target", "on Target", 
                                      "Above Target", "Total on Target and Above Target"),
                           ' ' = c(TOT, av.score, BT, Targ, AT, TTAT))  %>% 
    rename_with(~ c(' '), 2) %>% 
    flextable() %>% 
    font(part = "body", fontname = "Open Sans") %>% 
    font(part = "header", fontname = "Playfair Display") %>% 
    width(j = 1, "width"=75, unit = "mm") %>% 
    width(j = 2, "width"=20, unit = "mm") %>% 
    bold(part="body", i=6) %>% 
    bg(bg="#FECFB9", part="body", i=3) %>% 
    bg(bg="#FFFFAB", part="body", i=4) %>% 
    bg(bg="#E4F3D9", part="body", i=5) %>% 
    footnote(i = 1, j = 1, 
             value = as_paragraph("Total number of assessments used to evaluate the current Learning Outcome."), 
             part="body", ref_symbols = "1") %>% 
    footnote(i = 2, j = 1, 
             value = as_paragraph("Value on a scale of 0 to 20."), 
             part="body", ref_symbols = "2") %>% 
    fontsize(part = "footer", size = 8) %>% 
    font(part = "footer", fontname = "Open Sans")
  
  
  
  
}

# função para criar a classificação#

get_grade <- function(x) {
  grades <- ifelse(x > 90, "A", 
                   ifelse(x > 80, "B", 
                          ifelse(x > 60, "C","F")))
  return(grades)
}

# get_grade <- function(x) {
#   if (x > 90) {
#     return("A")
#   } else if (x > 80) {
#     return("B")
#   } else if (x > 60) {
#     return("C")
#   } else {
#     return("F")
#   }
# }

# tabela grande dos dados#

resultados <- full_data %>% 
  group_by(Program, Outcome, Course, Level, Assessment, Revaluation) %>% 
  summarise(avg = round(mean(Score), 2),
            count = n(),
            BT = sum(Score < 12),
            OT = sum(Score >= 12 & Score < 16),
            AT = sum(Score >= 16)) %>% 
  mutate(TTAT = round(((OT + AT) / count) * 100, 2)) %>% 
  mutate(Val = get_grade(TTAT))


KPI_table <- function(a){
  
  
  # loop <- c('I','II','IV','V')
  # 
  # p.name <- program.list$Program[program.list$Program_code == a]
  # 
  # small_border = fp_border(color="gray", width = 1)
  # no_border = fp_border(color="gray", width = 0)
  # 
  # 
  # for (x in loop) {
  #   
  #   Out <- outcomes.programs[outcomes.programs$Program == a & outcomes.programs$Outcome_Code == x, 3]
  #   
  #   TOT <- subset(full.data ,full.data$Program == a & full.data$Outcome == x) %>% 
  #     nrow()
  #   
  #   TTAT <- subset(full.data ,full.data$Program == a & full.data$Outcome == x) %>% 
  #     filter(Score >= 12) %>% 
  #     nrow()
  #   
  #   TTAT <- percent(TTAT/TOT, accuracy = .01)
  #   
  #   if (TTAT == '100.00%'){
  #     LO.score <- 'A'
  #   } else if(TTAT >= '90%'){
  #     LO.score <- 'A'
  #   } else if(TTAT >= '80%'){
  #     LO.score <- 'B'
  #   } else if(TTAT >= '60%'){
  #     LO.score <- 'C'
  #   } else {
  #     LO.score <- 'F'
  #   }
  #   
  #   assign(paste0(x), c(x, Out, TTAT, LO.score))
  #   
  #   
  # }
  
  # p.name <- program.list$Program[program.list$Program_code == a]

  small_border = fp_border(color="gray", width = 1)
  no_border = fp_border(color="gray", width = 0)
  
  KPI_program <- subset(resultados ,resultados$Program == a) %>%
                      ungroup() %>% 
                      select(c(2,12,13,6)) %>% 
                      rename(Outcome_Code = Outcome)
  
  subset(outcomes.programs, outcomes.programs$Program == a) %>%
    inner_join(KPI_program, by = "Outcome_Code") %>% 
    select(-1) %>%
    mutate(Revaluation = ifelse(Revaluation == 1, "X", ifelse(Revaluation == 0, "", Revaluation))) %>% 
    rename_with(~ c(' ', 'Outcome', 'TTAT', 'Score', 'Revaluation' )) %>% 
    flextable() %>% 
    #add_header_row(values = p.name, colwidths = 5) %>% 
    #add_header_row(values = "Assurance of Learning KPIs for the", colwidths = 5) %>% 
    font(part = "header", fontname = "Open Sans", i=1) %>% 
    font(part = "body", fontname = "Open Sans", j=c(2:5)) %>% 
    font(part = "body", fontname = "Playfair Display", j=1) %>% 
    border_inner_h(part="body", border = small_border ) %>% 
    border_inner_h(part="header", border = no_border ) %>%
    border_inner_v(part="body", border = small_border ) %>% 
    align(align = "center", part = "header") %>%
    align(align = "center", part = "body", j=c(1,3:5)) %>%
    bold(part="body", j = c(1,5)) %>% 
    fontsize(part = "body", size = 9, j=2) %>%  #aqui#
    width(j = 2, "width"=75, unit = "mm") %>% 
    bg(bg="#C4FF9E", part="body", j = 4, i = ~ Score == "A") %>% 
    bg(bg="#D2FFFA", part="body", j = 4, i = ~ Score == "B") %>% 
    bg(bg="#FEE8BC", part="body", j = 4, i = ~ Score == "C") %>% 
    bg(bg="#FFCCBD", part="body", j = 4, i = ~ Score == "F") %>% 
    #fontsize(part = "header", size = 14, i=c(1:2)) %>% 
    footnote(i = 1, j = 3, 
             value = as_paragraph("Total on Target and Above Target"), 
             part="header", ref_symbols = "1") %>% 
    fontsize(part = "footer", size = 8) %>% 
    font(part = "footer", fontname = "Open Sans")
  
}

# função criar gráficos employers #

employ_graf <- function(a){
  
  subset(employ, Program == a) %>% 
    pivot_longer(cols = c(4:8), names_to = "LO", values_to = "Score") %>% 
    mutate(LO = factor(LO, levels = c("V", "IV", "III", "II", "I"))) %>%
    ggplot(aes(x = Score, y = LO)) +
    geom_boxplot(alpha = 0.4, fill = "#34BBDA", size = 1.5, color = "#34BBDA", 
                 outlier.shape = 23, outlier.size = 3, outlier.alpha = 0.4, 
                 outlier.color = "#34BBDA", outlier.fill = "#34BBDA",
                 width = 0.75) +
    scale_x_continuous(limits = c(1, 6), breaks = c(1:6)) +
    geom_vline(xintercept = 4, color = "#12326E", linetype = 2, linewidth = 1.5) +
    geom_vline(xintercept = 5.5, color = "#12326E", linetype = 2, linewidth = 1.5) +
    geom_text(aes(x = 4.5, y = 3, label = "Target"), size = 6, color = "#464646", family = "Open Sans Light", angle = 90) +
    geom_text(aes(x = 5.7, y = 3, label = "Above Target"), size = 6, color = "#464646", family = "Open Sans Light", angle = 90) +
    geom_text(aes(x = 2.5, y = 3, label = "Below Target"), size = 6, color = "#464646", family = "Open Sans Light", angle = 90) +
    theme_minimal() +
    theme(text = element_text(size = 16, family = "Open Sans Light"),
          axis.text.y = element_text(size = 16, family = "Playfair Display", angle = 0),
          axis.title.y = element_blank())
  
}

# função criar gráficos alunos #

alumn_graf <- function(a){
  
  subset(students, Program == a) %>% 
    pivot_longer(cols = c(2:6), names_to = "LO", values_to = "Score") %>% 
    mutate(LO = factor(LO, levels = c("V", "IV", "III", "II", "I"))) %>%
    ggplot(aes(x = Score, y = LO)) +
    geom_boxplot(alpha = 0.4, fill = "#34BBDA", size = 1.5, color = "#34BBDA", 
                 outlier.shape = 23, outlier.size = 3, outlier.alpha = 0.4, 
                 outlier.color = "#34BBDA", outlier.fill = "#34BBDA",
                 width = 0.75) +
    scale_x_continuous(limits = c(1, 6), breaks = c(1:6)) +
    geom_vline(xintercept = 4, color = "#12326E", linetype = 2, linewidth = 1.5) +
    geom_vline(xintercept = 5.5, color = "#12326E", linetype = 2, linewidth = 1.5) +
    geom_text(aes(x = 4.5, y = 3, label = "Target"), size = 6, color = "#464646", family = "Open Sans Light", angle = 90) +
    geom_text(aes(x = 5.7, y = 3, label = "Above Target"), size = 6, color = "#464646", family = "Open Sans Light", angle = 90) +
    geom_text(aes(x = 2.5, y = 3, label = "Below Target"), size = 6, color = "#464646", family = "Open Sans Light", angle = 90) +
    theme_minimal() +
    theme(text = element_text(size = 16, family = "Open Sans Light"),
          axis.text.y = element_text(size = 16, family = "Playfair Display", angle = 0),
          axis.title.y = element_blank())
  
}

# criar resultados do survey alunos #

resultados.alumn <- students %>% 
  pivot_longer(cols = c(2:7), names_to = "LO", values_to = "Score") %>% 
  group_by(Program, LO) %>% 
  summarise(
    avg = round(mean(Score, na.rm = TRUE), 2),
    count = n(),
    BT = sum(Score < 4, na.rm = TRUE),
    OT = sum(Score >= 4 & Score < 5.5, na.rm = TRUE),
    AT = sum(Score >= 5.5, na.rm = TRUE)
  ) %>%
  mutate(TTAT = round(((OT + AT) / count) * 100, 2)) %>% 
  mutate(Val = get_grade(TTAT))

# criar resultados do survey empregadores #

resultados.employers <- employ %>% 
  select(c(2,4:9)) %>% 
  pivot_longer(cols = c(2:7), names_to = "LO", values_to = "Score") %>% 
  drop_na() %>% 
  group_by(Program, LO) %>% 
  summarise(avg = round(mean(Score), 2),
            count = n(),
            BT = sum(Score < 4),
            OT = sum(Score >= 4 & Score < 5.5),
            AT = sum(Score >= 5.5)) %>% 
  mutate(TTAT = round(((OT + AT) / count) * 100, 2)) %>% 
  mutate(Val = get_grade(TTAT))

# criar tabela resultados completos #

tabela.completa <- resultados %>% 
                      as.data.frame() %>% 
                      select(c(1,2,6,13)) %>% 
                      mutate(type = if_else(Revaluation == 1, "Revaluation", "Regular")) %>% 
                      select(-Revaluation) %>% 
                      rename(LO = Outcome)

tabela.completa <- resultados.alumn %>% 
  select(c(1,2,9)) %>% 
  mutate(type = "Alumni") %>% 
  bind_rows(tabela.completa)

tabela.completa <- resultados.employers %>% 
  select(c(1,2,9)) %>% 
  mutate(type = "Employer") %>% 
  bind_rows(tabela.completa)

tabela.completa <- AY2223 %>% 
  mutate(type = "AY2223") %>% 
  bind_rows(tabela.completa)

tabela.completa <- tabela.completa %>% 
  group_by(Program, LO) %>%
  pivot_wider(names_from = type, values_from = Val) %>% 
  mutate(Program = factor(Program, levels = program_levels),
         LO = factor(LO, levels = lo_levels)) %>% 
  as.data.frame()



tabela.completa <- tabela.completa[order(tabela.completa$Program, tabela.completa$LO), ]

tabela.completa <- tabela.completa[,c(1,2,3,6,7,4,5)]


# Função tabela completa por programa #

KPI_total_table <- function(a){
  
  # p.name <- program.list$Program[program.list$Program_code == a]
  
  small_border = fp_border(color="gray", width = 1)
  no_border = fp_border(color="gray", width = 0)
  
  subset(tabela.completa, Program == a) %>% 
    select(-1) %>% 
    rename_with(~ c(' ', 'AY2122', 'Reassessment', 'Regular', 'Employer', 'Alumni' )) %>%
    flextable() %>% 
    #add_header_row(values = p.name, colwidths = 6) %>% 
    #add_header_row(values = "Assurance of Learning KPIs for the", colwidths = 6) %>% 
    font(part = "header", fontname = "Open Sans", i=1) %>% 
    font(part = "body", fontname = "Open Sans", j=c(2:6)) %>% 
    font(part = "body", fontname = "Playfair Display", j=1) %>% 
    border_inner_h(part="body", border = small_border ) %>% 
    # border_inner_h(part="header", border = no_border ) %>%
    border_inner_v(part="body", border = small_border ) %>% 
    align(align = "center", part = "header") %>%
    align(align = "center", part = "body", j=c(1:6)) %>%
    bold(part="body", j = c(1)) %>% 
    fontsize(part = "body", size = 10, j=(2:6)) %>% 
    bg(bg="#C4FF9E", part="body", j=(2), i = ~ AY2122 == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(2), i = ~ AY2122 == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(2), i = ~ AY2122 == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(2), i = ~ AY2122 == "F") %>% 
    bg(bg="#C4FF9E", part="body", j=(3), i = ~ Reassessment == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(3), i = ~ Reassessment == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(3), i = ~ Reassessment == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(3), i = ~ Reassessment == "F") %>%
    bg(bg="#C4FF9E", part="body", j=(4), i = ~ Regular == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(4), i = ~ Regular == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(4), i = ~ Regular == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(4), i = ~ Regular == "F") %>%
    bg(bg="#C4FF9E", part="body", j=(5), i = ~ Employer == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(5), i = ~ Employer == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(5), i = ~ Employer == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(5), i = ~ Employer == "F") %>%
    bg(bg="#C4FF9E", part="body", j=(6), i = ~ Alumni == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(6), i = ~ Alumni == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(6), i = ~ Alumni == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(6), i = ~ Alumni == "F") %>% 
    width(j = c(2:6), "width"=25, unit = "mm") %>% 
    fontsize(part = "header", size = 9, i=1) %>% 
    width(j = 1, "width"=10, unit = "mm") %>% 
    footnote(i = 1, j = 2, 
             value = as_paragraph("Academic year 2021/2022"), 
             part="header", ref_symbols = "1") %>% 
    fontsize(part = "footer", size = 8) %>%
    line_spacing(part = "footer", space = 0.1) %>% 
    font(part = "footer", fontname = "Open Sans")
    
  
}

# Função tabela completa por programa - exportação para slide de powerpoint #

KPI_total_table_ppt <- function(a){
  
  p.name <- program.list$Program[program.list$Program_code == a]
  
  small_border = fp_border(color="gray", width = 1)
  no_border = fp_border(color="gray", width = 0)
  
  subset(tabela.completa, Program == a) %>% 
    select(-1) %>% 
    rename_with(~ c(' ', 'AY2223', 'Reassessment', 'Assessment', 'Employer', 'Alumni' )) %>%
    unnest(cols = c(' ', 'AY2223', 'Reassessment', 'Assessment', 'Employer', 'Alumni' )) %>% 
    flextable() %>% 
    font(part = "body", fontname = "Open Sans", j=c(2:6)) %>% 
    font(part = "body", fontname = "Playfair Display", j=1) %>%
    font(part = "header", fontname = "Playfair Display") %>%
    border_inner_h(part="body", border = small_border ) %>% 
    border_inner_h(part="header", border = no_border ) %>%
    border_inner_v(part="body", border = small_border ) %>% 
    align(align = "center", part = "body", j=c(1:6)) %>%
    align(align = "center", part = "header") %>% 
    bold(part="body", j = c(1)) %>% 
    fontsize(part = "body", size = 18, j=(1:6)) %>%
    fontsize(part = "header", size = 12, j= c(2:6)) %>%
    fontsize(part = "header", size = 16, j= 1) %>%
    bg(bg="#C4FF9E", part="body", j=(2), i = ~ AY2223 == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(2), i = ~ AY2223 == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(2), i = ~ AY2223 == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(2), i = ~ AY2223 == "F") %>% 
    bg(bg="#C4FF9E", part="body", j=(3), i = ~ Reassessment == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(3), i = ~ Reassessment == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(3), i = ~ Reassessment == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(3), i = ~ Reassessment == "F") %>%
    bg(bg="#C4FF9E", part="body", j=(4), i = ~ Assessment == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(4), i = ~ Assessment == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(4), i = ~ Assessment == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(4), i = ~ Assessment == "F") %>%
    bg(bg="#C4FF9E", part="body", j=(5), i = ~ Employer == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(5), i = ~ Employer == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(5), i = ~ Employer == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(5), i = ~ Employer == "F") %>%
    bg(bg="#C4FF9E", part="body", j=(6), i = ~ Alumni == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(6), i = ~ Alumni == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(6), i = ~ Alumni == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(6), i = ~ Alumni == "F") %>% 
    width(j = c(2:6), "width"=30, unit = "mm") %>% 
    save_as_image(path = paste0("KPIs totais/",a,".png"), width = 8, height = 6, res = 300)
    
  
}

print.program.b <- function(a){
  
  I(paste0("**",program.list$Program[program.list$Program_code == a],"**"))
}


print.outcome <- function(a,b){
  
  
  I(outcomes.programs$Outcome[outcomes.programs$Program == a & outcomes.programs$Outcome_Code == b])
  
}

print.course <- function(a,b,c){
  
  I(config$Course[config$Program == a & config$Outcome == b & config$Revaluation == c])
  
}

print.assessment <- function(a,b,c){
  
  I(config$Assessment[config$Program == a & config$Outcome == b & config$Revaluation == c])
  
}

# função para determinar o número de respondentes ao survey de alumni #

print.alumni.num <- function(a){
  
  I(paste0("**",as.integer(resultados.alumn$count[resultados.alumn$Program == a & resultados.alumn$LO == "I"]),"**"))
  
}

# função para determinar o número de respondentes ao survey de empregadores #

print.employers.num <- function(a){
  
  I(paste0("**",as.integer(resultados.employers$count[resultados.employers$Program == a & resultados.employers$LO == "I"]),"**"))
  
}

# função tabela resultados alumni #

Alumni_table <- function(a){
  
  LO_list <- subset(outcomes.programs, Program == a) %>% 
    select(-1) %>% 
    rename(LO = 1)
  
  
  small_border = fp_border(color="gray", width = 1)
  no_border = fp_border(color="gray", width = 0)
  
  subset(resultados.alumn, Program == a) %>% 
    inner_join(LO_list, by = "LO") %>% 
    ungroup() %>% 
    select(-Program,-c(4:7)) %>%
    select(c(1,5,2:4)) %>% 
    rename_with(~ c(' ', 'Outcome', 'Mean', 'TTAT', 'Score' )) %>%
    flextable() %>% 
      colformat_num(j = "Mean", digits = 2) %>% 
      colformat_num(j = "TTAT", digits = 2) %>% 
      bg(bg="#C4FF9E", part="body", j=(5), i = ~ Score == "A") %>% 
      bg(bg="#D2FFFA", part="body", j=(5), i = ~ Score == "B") %>% 
      bg(bg="#FEE8BC", part="body", j=(5), i = ~ Score == "C") %>% 
      bg(bg="#FFCCBD", part="body", j=(5), i = ~ Score == "F") %>% 
      font(part = "body", fontname = "Open Sans", j=c(2:5)) %>% 
      font(part = "body", fontname = "Playfair Display", j=1) %>% 
      border_inner_h(part="body", border = small_border ) %>% 
      border_inner_h(part="header", border = no_border ) %>%
      border_inner_v(part="body", border = small_border ) %>% 
      align(align = "center", part = "body", j=c(1,3:5)) %>%
      align(align = "center", part = "header", j=c(2:5)) %>% 
      bold(part="body", j = c(1)) %>% 
      fontsize(part = "body", size = 9, j=2) %>% 
      width(j = 2, "width"=75, unit = "mm") %>% 
      footnote(i = 1, j = 3, 
             value = as_paragraph("On a Likert scale of 1 to 6"), 
             part="header", ref_symbols = "1") %>% 
      footnote(i = 1, j = 4, 
             value = as_paragraph("Total on Target and Above Target"), 
             part="header", ref_symbols = "2") %>% 
      fontsize(part = "footer", size = 8) %>% 
      line_spacing(part = "footer", space = 0.1) %>% 
      font(part = "footer", fontname = "Open Sans")
      
  
}

# função tabela resultados alumni #

Employer_table <- function(a){
  
  LO_list <- subset(outcomes.programs, Program == a) %>% 
    select(-1) %>% 
    rename(LO = 1)
  
  
  small_border = fp_border(color="gray", width = 1)
  no_border = fp_border(color="gray", width = 0)
  
  subset(resultados.employers, Program == a) %>% 
    inner_join(LO_list, by = "LO") %>% 
    ungroup() %>% 
    select(-Program,-c(4:7)) %>%
    select(c(1,5,2:4)) %>% 
    rename_with(~ c(' ', 'Outcome', 'Mean', 'TTAT', 'Score' )) %>%
    flextable() %>% 
    colformat_num(j = "Mean", digits = 2) %>% 
    colformat_num(j = "TTAT", digits = 2) %>% 
    bg(bg="#C4FF9E", part="body", j=(5), i = ~ Score == "A") %>% 
    bg(bg="#D2FFFA", part="body", j=(5), i = ~ Score == "B") %>% 
    bg(bg="#FEE8BC", part="body", j=(5), i = ~ Score == "C") %>% 
    bg(bg="#FFCCBD", part="body", j=(5), i = ~ Score == "F") %>% 
    font(part = "body", fontname = "Open Sans", j=c(2:5)) %>% 
    font(part = "body", fontname = "Playfair Display", j=1) %>% 
    border_inner_h(part="body", border = small_border ) %>% 
    border_inner_h(part="header", border = no_border ) %>%
    border_inner_v(part="body", border = small_border ) %>% 
    align(align = "center", part = "body", j=c(1,3:5)) %>%
    align(align = "center", part = "header", j=c(2:5)) %>% 
    bold(part="body", j = c(1)) %>% 
    fontsize(part = "body", size = 9, j=2) %>% 
    width(j = 2, "width"=75, unit = "mm") %>% 
    footnote(i = 1, j = 3, 
             value = as_paragraph("On a Likert scale of 1 to 6"), 
             part="header", ref_symbols = "1") %>% 
    footnote(i = 1, j = 4, 
             value = as_paragraph("Total on Target and Above Target"), 
             part="header", ref_symbols = "2") %>% 
    fontsize(part = "footer", size = 8) %>% 
    line_spacing(part = "footer", space = 0.1) %>% 
    font(part = "footer", fontname = "Open Sans")
  
  
}

# exportação para excel de resultados #

# write.xlsx(resultados, "results/resultados.xlsx", row.names = FALSE)
# 
# write.xlsx(tabela.completa, "results/tabela.completa.xlsx", row.names = FALSE)

# compilação do relatório #

save.image(file="quarto/all_data.RData")

# quarto::quarto_render("quarto/report.qmd", output_format = "pdf")

