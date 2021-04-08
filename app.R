# Source script for app
source('./app_fn.R')
# Source script for uploading to server
source('./jvb.R')
# Source the password for the server
source('./pwd.R')

# Address for devwarehouse
#https://devwarehouse.indicia.org.uk/index.php/survey/edit/500

# Set folder name for emails
folderName = "Inbox"

## create outlook object
OutApp <- COMCreate("Outlook.Application")
outlookNameSpace = OutApp$GetNameSpace("MAPI")

# Create list of emails
folder <- outlookNameSpace$Folders(userName)$Folders(folderName)
emails <- folder$Items
num_emails <- folder$Items()$Count()

# Emails may not be in date order, so assign a lookup df to date order, but
# only if computerspeed = 1 (fast) or 2 (middling)
# This is to prevent this running on a slow computer
computerspeed <- 2
if(computerspeed == 1){
  cat('Reading in emails\n')
  datesDF <- pblapply(1:num_emails, FUN = function(i){
    dateEmail <- getDate(emails(i))
    data.frame(dates = as.Date(dateEmail),
               datetime = dateEmail,
               subj = emails(i)[['Subject']],
               Sender = getSender(emails(i)),
               j = i,
               stringsAsFactors = FALSE)
  }) %>% bind_rows()
  datesDF <- datesDF %>% arrange(desc(datetime))
} else if(computerspeed == 2){
  cat('Reading in email dates\n')
  datesDF <- pblapply(1:num_emails, FUN = function(i){
    dateEmail <- getDate(emails(i))
    data.frame(dates = as.Date(dateEmail),
               datetime = dateEmail,
               subj = '',
               Sender = '',
               j = i,
               stringsAsFactors = FALSE)
  }) %>% bind_rows()
  datesDF <- datesDF %>% arrange(desc(datetime))
} else {
  datesDF <- data.frame(dates = '',
                        subj = '',
                        Sender = '',
                        j = 1:num_emails)
}

datesDF$i <- 1:nrow(datesDF)
datesDF$Subject <- substr(datesDF$subj,1,100)
datesDF$Date <- as.character(datesDF$dates)

global <- list(
  sender = getSender(emails(datesDF$j[1])),
  sendername = emails(datesDF$j[1])[['SenderName']],
  subject = emails(datesDF$j[1])[['Subject']],
  msgbody = emails(datesDF$j[1])[['Body']],
  date = as.Date(getDate(emails(datesDF$j[1]))),
  datetime = getDate(emails(datesDF$j[1])),
  tel = '',
  location = '',
  comment = '',
  correspondance = '',
  body = '',
  geoparsed = data.frame(),
  num_attachments = emails(datesDF$j[1])[['attachments']]$Count(),
  attachment_location = 'tmp',
  expert = '',
  location_description = ''
)

# Define the UI
ui <- fluidPage(
  
  sidebarLayout(
    sidebarPanel(
      tabsetPanel(
    tabPanel('Upload',
      fluidRow(column(6,textInput(inputId = 'sender', label = 'Sender',
                                  value = global$sender)),
               column(6,textInput(inputId = 'name', label = 'Name',
                                  value = global$sendername))),
      textInput(inputId = 'subject', label = 'Subject',
                value = global$subject),
      fluidRow(column(5,
                      textInput(inputId = 'date', label = 'Date',
                                value = as.character(global$date))),
               column(7,
                      selectInput(inputId = 'species', label = 'Species',
                                  choices = c('Vespa velutina',''),
                                  selected = 'Vespa velutina'))),
      fluidRow(column(5,
                      textInput(inputId = 'location', label = 'Location',
                                placeholder = 'gridref of observation')),
               column(7,
                      textInput(inputId = 'tel', label = 'Telephone Number',
                                value = ''))),
      textInput(inputId = 'location_description',
                label = 'Location Description', value = ''),
      textInput(inputId = 'comment', label = 'Comment', value = ''),
      textAreaInput(inputId = 'correspondance', label = 'Correspondence',
                    height = '100px', value = global$msgbody),
      selectInput(inputId = 'expert', label = 'Expert Knowledge?',
                  choices = c('',
                              'General nature recording',
                              'Entomology',
                              'Apiculture'),
                  selected = ''),
      checkboxInput(inputId = 'includeAtt', label = 'Include Attachment Images',
                    value = TRUE),
      actionButton(inputId = 'upload_Indicia', label = 'Upload to Database'),
      textOutput('serverResponse'),
    ),
    tabPanel('Email',
             textInput(inputId = 'recipient', label = 'Recipient',
                       value = global$sender),
             textInput(inputId = 'subject_reply', label = 'Subject',
                       value = global$subject),
             selectInput(inputId = 'email_text_selector', label = 'Email Response',
                         choices = c(names(responses)),
                         selected = 'Custom'),
             textAreaInput(inputId = 'email_text',height = '100px',
                           label = 'Message Body', value = global$body),
             textOutput('sendemail'),
             actionButton(inputId = 'send_thanksbutno', label = 'Send reply'),
             checkboxInput(inputId = 'emailOn', label = 'Turn on Email Function',
                           value = FALSE),
             ),
    tabPanel('Jump',
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
    tabPanel('Tools',
             HTML('<br>'),
             fluidRow(
               column(6,
                      actionButton(inputId='launchBrowser',label='GridRef Finder')),
               column(6,
                      actionButton(inputId='launchBrowser2',label='GAGR'))),
             HTML("<hr>"),
             fluidRow(
               column(6,
                 bsButton(inputId = 'clearActions', label = 'Clear Actions')),
               column(6,
                 checkboxInput(inputId = 'turnonclearActions', label = 'Turn on Clear'))),
             textOutput('clearMessage'),
             HTML("<hr>"),
             fluidRow(
               column(6,
                      actionButton(inputId = 'geoparse', label = 'Attempt to Geoparse')),
               column(6,
                      actionButton(inputId = 'cleargeoparse', 'Clear Table'))),
             dataTableOutput(outputId = 'geotable')
             ))),
    mainPanel(fluidRow(column(2,bsButton(inputId = 'aftten', label = '',style = 'info',
                                         icon = icon('arrow-circle-left', 'fa-2x')) %>%
                         myPopify(txt = 'Go back 10 emails')),
                       column(2,bsButton(inputId = 'aft', label = '',style = 'primary',
                                         icon = icon('arrow-left', 'fa-2x')) %>%
                                myPopify(txt = 'Go back 1 email')),
                       column(2,bsButton(inputId = 'aft_img', label = '',style = 'success',
                                         icon = icon('chevron-left', 'fa-2x')) %>%
                                myPopify(txt = 'Go back 1 attachment')),
                       column(2,bsButton(inputId = 'fore_img', label = '',style = 'success',
                                         icon = icon('chevron-right', 'fa-2x')) %>%
                                myPopify(txt = 'Go forward 1 attachment')),
                       column(2,bsButton(inputId = 'fore',label = '',style = 'primary',
                                         icon = icon('arrow-right', 'fa-2x')) %>%
                                myPopify(txt = 'Go forward 1 email')),
                       column(2,bsButton(inputId = 'foreten', label = '',style = 'info',
                                         icon = icon('arrow-circle-right', 'fa-2x')) %>%
                                myPopify(txt = 'Go forward 10 emails'))
                       ),
              textOutput('attachment_info'),
              actionButton("att_open", "Open File"),
              imageOutput('myImage', height = '100%'),
              htmlOutput('msgbody'),
              htmlOutput("inc")
    )
  )
)

# Create the server
server <- function(input, output, session){

  values <-
    reactiveValues(i = 1,
      sender = getSender(datesDF$j[1]),
      sendername = emails(datesDF$j[1])[['SenderName']],
      subject = emails(datesDF$j[1])[['Subject']],
      msgbody = emails(datesDF$j[1])[['Body']],
      date = as.Date(getDate(emails(datesDF$j[1]))),
      datetime = getDate(emails(datesDF$j[1])),
      attachments = ifelse(emails(datesDF$j[1])[['attachments']]$Count()>0,
                           emails(datesDF$j[1])[['attachments']]$Item(1)[['DisplayName']],
                           ''),
      num_attachments = emails(datesDF$j[1])[['attachments']]$Count(),
      attachment_location = 'tmp',
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
  observeEvent(input$foreten, {
    values$i <- values$i - 10
    if(values$i<0){
      values$i <- 1
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
  observeEvent(input$aftten, {
    values$i <- values$i + 10
    if(values$i>values$num_emails){
        values$i <- values$num_emails
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
      global$attachment_location <- values$attachment_location
      #cat(global$attachment_location,'\n')
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
      global$attachment_location <- values$attachment_location
      #cat(global$attachment_location,'\n')
    }
  })
  
  output$msgbody <- renderUI({
    HTML(paste(str_replace_all(values$msgbody,'\\n','<br>')))
  })
  
  if(computerspeed == 1){
    output$summaryDF <- renderDataTable({
      (datesDF %>% select(i, Subject, datetime, Sender))
    })
  }
  
  observeEvent(input$att_open, {
    if(values$num_attachments>0){
      shell(paste0(tmpfiles, '/', values$attachment_location))
    }
  })

  # Attempt to geoparse
  observeEvent(input$geoparse, {
    output$geotable <- renderDataTable({
      # Check if there's a postcode in the text
      pcdf <- getpostcode(values$msgbody)
      
      # Geoparse text with a progress bar
      withProgress(message = 'Geoparsing...', value = 0, {
        # Take text and split by words, removing stopwords
        textl <- gsub('[[:punct:] ]+',' ',values$msgbody) %>% tolower()
        textl <- strsplit(values$msgbody,split = ' ') %>% unlist()
        textl <- textl[!duplicated(textl)]
        textl <- textl[!(textl %in% stopwrds)]
        
        # Try to find words in geonames
        results <- lapply(1:length(textl), FUN = function(l){
          incProgress(1/length(textl))
          geoparse(textl, l)
        }) %>% bind_rows()
        names(results) <- c('lat','lng','name')
      })
      
      global$geoparsed <- bind_rows(pcdf, results)
      global$geoparsed
    })
  })
  
  observeEvent(input$cleargeoparse, {
    global$geoparsed <- data.frame()
    output$geotable <- renderDataTable({
      global$geoparsed
    })
  })
  
  observeEvent(input$email_text_selector, {
    global$body <- responses[[input$email_text_selector]]
    updateTextAreaInput(session, inputId = 'email_text',
                        label = 'Message Body', value = global$body)
  })
  
  # Clear Actions rds
  observeEvent(input$clearActions, {
    if(input$turnonclearActions){
      overwrite_actions()
      output$clearMessage <- renderText({
        paste0('Actions File Cleared')
      })
    } else {
      output$clearMessage <- renderText({
        paste0('Actions File NOT Cleared - please check the \'Turn on Clear\' button')
      })
    }
  })

  # Send an email if this button is pressed
  observeEvent(input$send_thanksbutno, {
    if(input$emailOn){
      if(!grepl(pattern = "^[[:alnum:].-_]+@[[:alnum:].-]+$",
                x = input$recipient)){
        output$sendemail <- renderText({
          paste0('Please enter a valid email address and try again')
        })
      } else {
        values <-
          send_email(OutApp = OutApp,
                     values = values,
                     reply = input$email_text,
                     recipient = input$recipient,
                     msgBody = input$email_text,
                     subject = input$subject_reply,
                     from = from)
        output$sendemail <- renderText({
          paste0('Email sent')
        })
        updateactions(currentemail = emails(datesDF$j[values$i]),
                      action = 'reply')
      }
    } else {
      output$sendemail <- renderText({
        paste0('Email not sent - please check the \'Turn on Email Function\' button')
      })
    }
  })

  # Turn on attachment flag if ticked
  observeEvent(input$includeAtt,{
    values$includeAtt <- input$includeAtt
  })
  
  observeEvent(input$launchBrowser,{
    output$inc <- renderUI({
      getPage('https://gridreferencefinder.com/')
    })
  })
  
  observeEvent(input$launchBrowser2,{
    output$inc <- renderUI({
      getPage('https://www.bnhs.co.uk/2019/technology/grabagridref/gagrol.php#map')
    })
  })

  # Upload the record to Indicia
  observeEvent(input$upload_Indicia, {
    if(input$tel == ''){
      global$tel <- NULL
    } else {
      global$tel <- input$tel
    }
    global$location <- input$location
    global$correspondance <- input$correspondance
    global$comment <- input$comment
    global$expert <- input$expert
    global$location_description <- input$location_description
    global$sender <- input$sender
    global$date <- input$date
    global$sendername <- input$name
    imageStr <- NULL
    if(values$includeAtt){
      # Attachment images are being included in the upload.
      #  Find out what they are and store temporary copies
      imagelist <- getallimages(emails, values, datesDF)
      if(!is.null(imagelist)){
        cat('\nUploading images to data warehouse\n')
        imageStr <- pblapply(imagelist, FUN = function(img){
          getnonce(password = password, URLbase = URLbase) %>%
            postimage(imgpath = img, URLbase = URLbase)
        }) %>% unlist()
      }
    }
    cat('\nUploading record to data warehouse\n')
    submission <- createjson(imgString = imageStr,
                             email = global$sender,
                             recordername = global$sendername,
                             tel = global$tel,
                             date = global$date,
                             location = global$location,
                             comment = global$comment,
                             correspondance = global$correspondance,
                             experience = global$expert,
                             location_description = global$location_description)
    if(submission=='Location improperly formatted'){
      output$serverResponse <- renderText({
        paste0(submission)
      })
    } else {
      serverPost <- getnonce(password = password, URLbase = URLbase) %>%
        postsubmission(URLbase = URLbase,
                       submission = submission)
      serverOut <- serverPost %>% fromJSON()
      serverResp <- paste0('SUCCESS! ',
                           'Sample ID: ',serverOut$outer_id,
                           ', Occurrence ID: ',
                           serverOut$struct$children %>%
                             filter(model == 'occurrence') %>% pull(id))
      cat(serverResp,'\n')
      
      output$serverResponse <- renderText({
        paste0(serverResp)
      })
      updateactions(currentemail = emails(datesDF$j[values$i]),
                    action = 'upload',
                    sampleID = serverOut$outer_id,
                    occurrenceID = serverOut$struct$children %>%
                      filter(model == 'occurrence') %>% pull(id))
    }
  })
}

shinyApp(ui = ui, server = server)
