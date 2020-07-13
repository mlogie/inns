# Source script for app
source('./app_fn.R')
# Source script for uploading to server
source('./jvb.R')
# Source the password for the server
source('./pwd.R')

# Set folder name for emails
folderName = "Inbox"

## create outlook object
OutApp <- COMCreate("Outlook.Application")
outlookNameSpace = OutApp$GetNameSpace("MAPI")

# Some code to search - saving for later development
#search.phrase <- '2020-07-09'
#search <- OutApp$AdvancedSearch(
#  "Inbox",
#  paste0("http://schemas.microsoft.com/mapi/proptag/0x0037001E ci_phrasematch '", search.phrase, "'")
#)
#results <- search[['Results']]
#results[[1]][['Subject']]
#results$Count()

# Create list of emails
folder <- outlookNameSpace$Folders(1)$Folders(folderName)
emails <- folder$Items
num_emails <- folder$Items()$Count()

# Emails may not be in date order, so assign a lookup df to date order, but
# only if computerspeed = 1 (fast) or 2 (middling)
# This is to prevent this running on a slow computer
computerspeed <- 2
if(computerspeed == 1){
  cat('Reading in emails\n')
  datesDF <- pblapply(1:num_emails, FUN = function(i){
    data.frame(dates = getDate(emails(i)),
               subj = emails(i)[['Subject']],
               Sender = getSender(emails(i)),
               j = i)
  }) %>% bind_rows()
  datesDF <- datesDF %>% arrange(desc(dates))
} else if(computerspeed == 2){
  cat('Reading in email dates\n')
  datesDF <- pblapply(1:num_emails, FUN = function(i){
    data.frame(dates = getDate(emails(i)),
               subj = '',
               Sender = '',
               j = i)
  }) %>% bind_rows()
  datesDF <- datesDF %>% arrange(desc(dates))
} else {
  datesDF <- data.frame(dates = '',
                        subj = '',
                        Sender = '',
                        j = 1:num_emails)
}

datesDF$i <- 1:nrow(datesDF)
datesDF$Subject <- substr(datesDF$subj,1,100)
datesDF$Date <- as.character(datesDF$dates)

global = list(
  sender = getSender(emails(datesDF$j[1])),
  subject = emails(datesDF$j[1])[['Subject']],
  msgbody = emails(datesDF$j[1])[['Body']],
  date = getDate(emails(datesDF$j[1])),
  tel = '',
  location = '',
  comment = '',
  correspondance = ''
)

# Define the UI
ui <- fluidPage(
  
  titlePanel('Invasive Species Alerts Tool'),
  
  sidebarLayout(
    sidebarPanel(
      fluidRow(
        column(6,
               actionButton(inputId = 'aft', label = 'Previous Email'),
               actionButton(inputId = 'aftten', label = 'Jump 10 Back'),
               HTML("<br><br>"),
               actionButton(inputId = 'aft_img', label = 'Previous Image')),
        column(6,
               actionButton(inputId = 'fore',label = 'Next Email'),
               actionButton(inputId = 'foreten', label = 'Jump 10 Forwards'),
               HTML("<br><br>"),
               actionButton(inputId = 'fore_img', label = 'Next Image'))
      ),
      HTML("<hr>"),
      textOutput('attachment_info'),
      HTML("<br>"),
      textInput(inputId = 'sender', label = 'Sender',
                value = global$sender),
      textInput(inputId = 'subject', label = 'Subject',
                value = global$subject),
      textInput(inputId = 'date', label = 'Date',
                value = as.character(global$date)),
      textInput(inputId = 'species', label = 'Species',
                value = 'Vespa velutina'),
      textInput(inputId = 'location', label = 'Location',
                placeholder = 'gridref of observation'),
      textInput(inputId = 'tel', label = 'Telephone Number', value = ''),
      textInput(inputId = 'comment', label = 'Comment', value = ''),
      textInput(inputId = 'correspondance', label = 'Correspondence',
                value = ''),
      checkboxInput(inputId = 'includeAtt', label = 'Include Attachment Images',
                    value = TRUE),
      actionButton(inputId = 'upload_Indicia', label = 'Upload to Database'),
      textOutput('serverResponse'),
      HTML("<br><br>"),
      actionButton(inputId = 'send_thanksbutno', label = 'Send reply'),
      HTML("<hr>"),
      if(computerspeed <= 2){
        dateInput(inputId = 'dateselector', label = 'Select Email Date')
      },
      fluidRow(
        column(6,
               textInput(inputId = 'i', label = 'Select Index (i)',
                         value = '1')),
        column(6,
               HTML("<br>"),
               actionButton(inputId = 'jumpToIndex', label = 'Jump to Index'))
        ),
      HTML("<hr>"),
      if(computerspeed == 1){
        dataTableOutput(outputId = 'summaryDF')
      }
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
      date = getDate(emails(datesDF$j[1])),
      attachments = ifelse(emails(datesDF$j[1])[['attachments']]$Count()>0,
                           emails(datesDF$j[1])[['attachments']]$Item(1)[['DisplayName']],
                           ''),
      num_attachments = emails(datesDF$j[1])[['attachments']]$Count(),
      num_emails = num_emails,
      img_num = 1,
      includeAtt = TRUE)
  
  # Jump to selected date
  observeEvent(input$dateselector, {
    if(computerspeed <= 2){
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
    }
  })
  
  # Jump to selected index value
  observeEvent(input$jumpToIndex, {
    if(any(as.character(datesDF$i)==input$i) & !is.na(as.numeric(input$i))){
      values$i <- as.numeric(input$i)
    } else {
      values$i <- 1
    }

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
  
  # Going forward in time, subtract ten from the email counter (i),
  # or loop to the end if we hit the beginning
  observeEvent(input$foreten, {
    values$i <- values$i - 10
    if(values$i<0){
      values$i <- values$i + values$num_emails
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
  
  # Going backward in time, add ten to the email counter (i),
  # or loop back to the beginning if we hit the end
  observeEvent(input$aftten, {
    values$i <- values$i + 10
    if(values$i>values$num_emails){
        values$i <- values$i-values$num_emails
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
  
  output$msgbody <- renderText({
    paste(values$msgbody)
  })
  
  if(computerspeed == 1){
    output$summaryDF <- renderDataTable({
      (datesDF %>% select(i, Subject, Date, Sender))
    })
  }

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
  
  # Turn on attachment flag if ticked
  observeEvent(input$includeAtt,{
    values$includeAtt <- input$includeAtt
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
    global$correspondance <- input$correspondance
    global$comment <- input$comment
    imageStr <- NULL
    if(values$includeAtt){
      # Attachment images are being included in the upload.
      #  Find out what they are and store temporary copies
      imagelist <- getallimages(emails, values, datesDF)
      if(!is.null(imagelist)){
        cat('\nUploading images to data warehouse\n')
        imageStr <- pblapply(imagelist, FUN = function(img){
          getnonce(password = password) %>%
            postimage(imgpath = img)
        }) %>% unlist()
      }
    }
    cat('\nUploading record to data warehouse\n')
    serverPost <- getnonce(password = password) %>%
      postsubmission(submission = createjson(imgString = imageStr,
                                             email = global$sender,
                                             tel = global$tel,
                                             date = global$date,
                                             location = global$location,
                                             comment = global$comment,
                                             correspondance = global$correspondance))
    serverOut <- serverPost %>% fromJSON()
    cat('Done\n')
    
    output$serverResponse <- renderText({
      paste0('SUCCESS! ',
             'Sample ID: ',serverOut$outer_id,
             ', Occurrence ID: ',
             serverOut$struct$children %>%
               filter(model == 'occurrence') %>% pull(id))
    })
  })
}

shinyApp(ui = ui, server = server)
