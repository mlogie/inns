library(shiny)
library(imager)
library(jpeg)
library(png)
library(shinyFiles)
library(shinyWidgets)
library(sendmailR)
library(RDCOMClient)
source('./app_fn.R')
library(httr)
library(digest)
library(jsonlite)
library(dplyr)
library(pbapply)
source('./jvb.R')
source('./pwd.R')
#install.packages("RDCOMClient", repos = "http://www.omegahat.net/R")

# Set folder name for emails
folderName = "Inbox"

## create outlook object
OutApp <- COMCreate("Outlook.Application")
outlookNameSpace = OutApp$GetNameSpace("MAPI")

# Create list of emails
folder <- outlookNameSpace$Folders(1)$Folders(folderName)
emails <- folder$Items
num_emails <- folder$Items()$Count()

# Emails may not be in date order, so assign a lookup df to date order
cat('Reading in emails\n')
datesDF <- pblapply(1:num_emails, FUN = function(i){
  data.frame(dates = floor(emails(i)$ReceivedTime()) + as.Date("1899-12-30"),
             subj = emails(i)[['Subject']],
             Sender = getSender(emails(i)))
}) %>% bind_rows()
datesDF$j <- 1:nrow(datesDF)
datesDF <- datesDF %>% arrange(desc(dates))
datesDF$i <- 1:nrow(datesDF)
datesDF$Subject <- substr(datesDF$subj,1,15)
datesDF$Date <- as.character(datesDF$dates)

global = list(
  sender = getSender(emails(datesDF$j[1])),
  subject = emails(datesDF$j[1])[['Subject']],
  msgbody = emails(datesDF$j[1])[['Body']],
  date = as.Date("1899-12-30") + floor(emails(datesDF$j[1])$ReceivedTime()),
  tel = '',
  location = ''
)

# Define the UI
ui <- fluidPage(
  
  titlePanel('Invasive Species Alerts Tool'),
  
  sidebarLayout(
    sidebarPanel(
      actionButton(inputId = 'aft', label = 'Previous Email'),
      actionButton(inputId = 'fore',label = 'Next Email'),
      actionButton(inputId = 'aft_img', label = 'Previous Image'),
      actionButton(inputId = 'fore_img', label = 'Next Image'),
      dateInput(inputId = 'dateselector', label = 'Select Email Date'),
      htmlOutput('newline'),
      textOutput('sender'),
      textOutput('subject'),
      textOutput('date'),
      textOutput('attachment_info'),
      actionButton(inputId = 'send_thanksbutno', label = 'Send reply'),
      textInput(inputId = 'sender', label = 'Sender',
                value = global$sender),
      textInput(inputId = 'species', label = 'Species',
                value = 'Vespa velutina'),
      textInput(inputId = 'date', label = 'Date',
                value = as.character(global$date)),
      textInput(inputId = 'location', label = 'Location',
                placeholder = 'gridref of observation'),
      textInput(inputId = 'tel', label = 'Telephone Number', value = ''),
      actionButton(inputId = 'upload_Indicia', label = 'Upload to Database'),
      textInput(inputId = 'i', label = 'Enter Index',
                value = '1'),
      actionButton(inputId = 'jumpToIndex', label = 'Jump to Index'),
      dataTableOutput(outputId = 'summaryDF')
    ),
    
    mainPanel(
      imageOutput('myImage'),
      verbatimTextOutput('msgbody')
    )
  )
)

# Create the server
server <- function(input, output, session){

  values <-
    reactiveValues(i = 1,
      sender = getSender(datesDF$j[1]),
      subject = emails(datesDF$j[1])[['Subject']],
      msgbody = emails(datesDF$j[1])[['Body']],
      date = as.Date("1899-12-30") + floor(emails(datesDF$j[1])$ReceivedTime()),
      attachments = ifelse(emails(datesDF$j[1])[['attachments']]$Count()>0,
                           emails(datesDF$j[1])[['attachments']]$Item(1)[['DisplayName']],
                           ''),
      num_attachments = emails(datesDF$j[1])[['attachments']]$Count(),
      num_emails = num_emails,
      img_num = 1)
  
  output$newline <- renderUI({
    HTML(paste('<br>','','</br>'))
  })

  # Jump to selected date
  observeEvent(input$dateselector, {
    if(any(datesDF$dates==input$dateselector)){
      lastmatch <- which(datesDF$dates==input$dateselector) %>% tail(1)
      values$i <- datesDF$i[lastmatch]
    } else {
      # We don't have an email which matches that date, so find the nearest,
      # looking forward in time first
      diffs <- datesDF$dates - input$dateselector
      if(any(diffs > 0)){
        lastmatch <- which(diffs==min(diffs[diffs > 0])) %>% tail(1)
        values$i <- datesDF$i[lastmatch]
      } else {
        # A time further in the future than any emails has been picked.
        # Go to the top
        values$i <- 1
      }
    }
    
    # Call the wrapper function to jump to the email and get outputs
    retList <- jumpTo(emails, values, global, datesDF, output, session)
    output <- retList$output
    values <- retList$values
    global <- retList$global
  })
  
  # Jump to selected index value
  observeEvent(input$jumpToIndex, {
    if(any(as.character(datesDF$i)==input$i) & !is.na(as.numeric(input$i))){
      values$i <- as.numeric(input$i)
    } else {
      values$i <- 1
    }
    as.numeric(as.character('t'))
    # Call the wrapper function to jump to the email and get outputs
    retList <- jumpTo(emails, values, global, datesDF, output, session)
    output <- retList$output
    values <- retList$values
    global <- retList$global
  })
  
  # Going forward in time, subtract one from the email counter (i),
  # or loop to the end if we hit the beginning
  observeEvent(input$fore, {
    if(values$i!=1){
      values$i <- values$i - 1
    } else {
      values$i <- values$num_emails
    }
    
    # Call the wrapper function to jump to the email and get outputs
    retList <- jumpTo(emails, values, global, datesDF, output, session)
    output <- retList$output
    values <- retList$values
    global <- retList$global
  })
  
  # Going backward in time, add one to the email counter (i),
  # or loop back to the beginning if we hit the end
  observeEvent(input$aft, {
    if(values$i<values$num_emails){
      values$i <- values$i + 1
    } else {
      values$i <- 1
    }
    
    # Call the wrapper function to jump to an email and get outputs
    retList <- jumpTo(emails, values, global, datesDF, output, session)
    output <- retList$output
    values <- retList$values
    global <- retList$global
  })
  
  # Go backward one image in the email
  observeEvent(input$aft_img, {
    if(values$img_num != 1){
      values$img_num <- values$img_num - 1
    } else {
      values$img_num <- values$num_attachments
    }
    
    # No need to get any email info, just grab the relevant attachment
    if(values$num_attachments > 1)
    {
      return_list <- format_attachments(emails, values, output, datesDF)
      output <- return_list$output
      values <- return_list$values
    }
  })
  
  # Go forward one image in the email
  observeEvent(input$fore_img, {
    if(values$img_num < values$num_attachments){
      values$img_num <- values$img_num + 1
    } else {
      values$img_num <- 1
    }
    
    # No need to get any email info, just grab the relevant attachment
    if(values$num_attachments>1)
    {
      return_list <- format_attachments(emails, values, output, datesDF)
      output <- return_list$output
      values <- return_list$values
    }
  })
  
  output$subject <- renderText({
    paste(values$subject)
  })
  
  output$msgbody <- renderText({
    paste(values$msgbody)
  })
  
  output$date <- renderText({
    paste(values$date)
  })
  
  output$sender <- renderText({
    paste(values$sender)
  })
  
  output$summaryDF <- renderDataTable({
    (datesDF %>% select(i, Subject, Date, Sender))
  })

  # Send an email if this button is pressed
  observeEvent(input$send_thanksbutno, {
    values <-
      send_email(OutApp,
                 values,
                 reply =
        paste0('\r\n\r\nThis is not an Asian Hornet',
               ifelse(values$species=='','',paste0(', it is actually a ',values$species)),
               '.\r\n\r\nKeep up the good work.',
               '\r\n\r\nFrom GB Non-Native Species Information Portal (GB-NNSIP)'))
  })
  
  global$location <- reactive ({
    input$location
  })
  
  global$tel <- reactive ({
    input$tel
  })
  
  # Upload the record to Indicia
  observeEvent(input$upload_Indicia, {
      # json_sample <- getnonce(password = PASSWORD) %>%
      #   postimage(imgpath = 'test.png') %>% createjson(correspondance = 'test')
      #   #### Post your submission
      # serverPost <- getnonce(password = PASSWORD) %>%
      #   postsubmission(submission = json_sample)
    if(input$tel == ''){
      global$tel <- NULL
    } else {
      global$tel <- input$tel
    }
    global$location <- input$location
    serverPost <- getnonce(password = password) %>%
      postsubmission(submission = createjson(email = global$sender,
                                             tel = global$tel,
                                             date = global$date,
                                             location = global$location,
                                             correspondance = 'test'))
    cat(serverPost)
  })
  
  
  
}

shinyApp(ui = ui, server = server)
