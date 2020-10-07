library(shiny)
library(imager)
library(jpeg)
library(png)
library(shinyFiles)
library(shinyWidgets)
library(sendmailR)
#install.packages("RDCOMClient", repos = "http://www.omegahat.net/R")
library(RDCOMClient)
library(httr)
library(digest)
library(jsonlite)
library(dplyr)
library(pbapply)
library(tools)
library(tm)
library(geonames)
library(stringr)
library(shinyBS)
source('./responses.R')
tmpfiles <- dirname(tempfile())
addResourcePath('tmpfiles', tmpfiles)

# Function to create a shiny ready jpg object from a jpg file passed
# Function requires:
#   attachment_file - a file created with tempfile() then with
#     an attachment image saved to this temporary file
#   Optionally:
#     width, height - dimensions
#     alt - alt text for image
# Returns: list containing the image, which the renderImage shiny function understands
createJPEG <- function(attachment_file, width = 400, height = 400, alt = 'image'){
  myJPG <- readJPEG(attachment_file,native=TRUE)
  res = dim(myJPG)[2:1]
  tryCatch({
    plot(1,1,xlim=c(1,(res[1])),ylim=c(1,(res[2])),asp=1,type='n',
         xaxs='i',yaxs='i',xaxt='n',yaxt='n',xlab='',ylab='',bty='n')
  }, error = function(error) {
    plot(1,1,xlim=c(1,(res[2])),ylim=c(1,(res[1])),asp=1,type='n',
         xaxs='i',yaxs='i',xaxt='n',yaxt='n',xlab='',ylab='',bty='n')
  })
  rasterImage(myJPG,1,1,res[1],res[2])
  dev.off()
  list(src = attachment_file,
       contentType = 'image/jpeg',
       width = width,
       height = height,
       alt = alt)
}

# Function to create a shiny ready png object from a png file passed
# Function requires:
#   attachment_file - a file created with tempfile() then with
#     an attachment image saved to this temporary file
#   Optionally:
#     width, height - dimensions
#     alt - alt text for image
# Returns: list containing the image, which the renderImage shiny function understands
createPNG <- function(attachment_file, width = 400, height = 400, alt = 'image'){
  myPNG <- readPNG(attachment_file)
  plot.new()
  rasterImage(myPNG,0,0,1,1)
  dev.off()
  list(src = attachment_file,
       contentType = 'image/png',
       width = width,
       height = height,
       alt = alt)
}

# Function to get email sender from email pointer
getSender <- function(email){
  sender <- tryCatch({
    if(email$Sender()$AddressEntryUserType() == 30){
      email[['SenderEmailAddress']]
    } else {
      email[['Sender']][['GetExchangeUser']][['PrimarySmtpAddress']]
    }
  }, error = function(error) {
    return('Could not obtain sender')
  })
  if(is.null(sender)){
    sender <- 'Could not obtain sender'
  }
  sender
}

# Function to get date from email pointer
getDate <- function(email){
  dateOut <- tryCatch({
    floor(email$ReceivedTime()) + as.Date("1899-12-30")
  }, error = function(error) {
    return('Could not extract date')
  })
  dateOut
}

# Function to grab the 
# Function requires:
#   emails - RDCOM pointer to list of emails
#   values - shiny reactive values list
#   global - global variables for use in the UI
#   datesDF - dataframe of dates so emails come out in order
# Function returns:
#   a list of list(values = values, global = global)
# This function updates:
#   values/global$subject - subject
#   values/global$sender - sender
#   values/global$msgbody - message body
# The email for which information extracted is the value saved as values$i
extract_contents <- function(emails, values, global, datesDF){
  values$sender <- global$sender <- getSender(emails(datesDF$j[values$i]))
  values$sendername <- global$sendername <- emails(datesDF$j[values$i])[['SenderName']]
  values$subject <- global$subject <- emails(datesDF$j[values$i])[['Subject']]
  values$msgbody <- global$msgbody <- emails(datesDF$j[values$i])[['Body']]
  values$date <- global$date <- getDate(emails(datesDF$j[values$i]))
  values$num_attachments <- global$num_attachments <-
    emails(datesDF$j[values$i])[['attachments']]$Count()
  list(values = values, global = global)
}

# Function which grabs an attachment from an email, saving the JPEG/PNG attachment
# to output$myImage as a renderImage shiny object
# Function requires:
#   emails - RDCOM pointer to list of emails
#   values - shiny reactive values list
#   output - shiny output list
#   datesDF - dataframe of dates so emails come out in order
# Function returns a list containing: (output, values)
# This function updates:
#   values$attachments - the name of the current attachment (based on values$img_num)
#   values$num_attachments - the number of attachments
#   output$attachment_info - shiny renderText with info about the attachment
#     Options are:
#       "Attachment x of y"
#       "Unknown image format for attachment"
#       "No attachments"
format_attachments <- function(emails, values, output, datesDF){
  attach_obj <- emails(datesDF$j[values$i])[['attachments']]
  if(attach_obj$Count() > 0){
    values$attachments <- attach_obj$Item(values$img_num)[['DisplayName']]
    values$num_attachments <- attach_obj$Count()
    output$attachment_info <- renderText({
      paste0('Attachment ',values$img_num,' of ',values$num_attachments,
             ': ',values$attachments)
    })
    
    attachment_file <- tempfile()
    attach_obj$Item(values$img_num)$SaveAsFile(attachment_file)
    dispName <- attach_obj$Item(values$img_num)[['DisplayName']]
    newName  <- file.path(dirname(attachment_file),
                          dispName)
    file.rename(attachment_file,newName)
    values$attachment_location <- basename(newName)
    
    if(grepl('jpg$|jpeg$',tolower(values$attachments))){
      img <- readJPEG(newName)
      wh <- dim(img)
      if(wh[1]>400){
        wh[2] <- 400*wh[2]/wh[1]
        wh[1] <- 400
      }
      output$myImage <- renderImage({
        createJPEG(newName, width = wh[2], height = wh[1])
      }, deleteFile = FALSE)
      output$attachment_info <- renderText({
        paste0('File saved here:',newName)
      })
      
    } else if(grepl('png$',tolower(values$attachments))){
      img <- readPNG(newName)
      wh <- dim(img)
      if(wh[1]>400){
        wh[2] <- 400*wh[2]/wh[1]
        wh[1] <- 400
      }
      output$myImage <- renderImage({
        createPNG(newName, width = wh[2], height = wh[1])
      }, deleteFile = FALSE)
      output$attachment_info <- renderText({
        paste0('File saved here:',newName)
      })
      
    } else {
      output$myImage <- renderImage({
        createPNG('www/unknown_format.png', height = 50, width = 90)
      }, deleteFile = FALSE)
      
      output$attachment_info <- renderText({
        paste0('Unknown format for attachment: ',
               values$attachments,', ',
               values$img_num,' of ',values$num_attachments,
               '.\nFile saved here: ',newName)
      })
    }
  } else {
    values$num_attachments <- 0
    values$attachments <- ''
    output$myImage <- renderImage({
      createPNG('www/no_attachments.png', height = 50, width = 120)
    }, deleteFile = FALSE)
    
    output$attachment_info <- renderText({
      paste0('No attachments')
    })
  }
  list(output = output, values = values)
}

# Function getallimages.  Takes an email with attachments and stores them
#  locally in order to upload them to the data warehouse.  It will only do
#  this with attachments ending in .jpg, .jpeg, .png and .bmp
# Function requires:
#   emails - RDCOM pointer to list of emails
#   values - shiny reactive values list
#   datesDF - dataframe of dates so emails come out in order
# Function returns:
#   a vector of the absolute locations of the files
getallimages <- function(emails, values, datesDF){
  attach_obj <- emails(datesDF$j[values$i])[['attachments']]
  imagelist <- NULL
  if(attach_obj$Count()>0){
    imagelist <- lapply(1:attach_obj$Count(), FUN = function(k){
      dispName <- attach_obj$Item(k)[['DisplayName']]
      if(grepl('jpg$|png$|bmp$|jpeg$',dispName)){
        attachment_file <- tempfile()
        attach_obj$Item(k)$SaveAsFile(attachment_file)
        newName <- file.path(dirname(attachment_file),
                             dispName)
        file.rename(attachment_file,newName)
        return(newName)
      } else {
        return(NULL)
      }
    }) %>% unlist()
  }
  imagelist
}

# General wrapper function to jump to email i
jumpTo <- function(emails, values, global, datesDF, output, session){
  ecOut <- extract_contents(emails, values, global, datesDF)
  values <- ecOut$values
  global <- ecOut$global
  values$img_num <- 1
  # Grab the attachments
  return_list <- format_attachments(emails, values, output, datesDF)
  updateTextInput(session, inputId = 'sender', label = 'Sender',
                  value = global$sender)
  updateTextInput(session, inputId = 'name', label = 'Name',
                  value = global$sendername)
  updateTextInput(session, inputId = 'recipient', label = 'Recipient',
                  value = global$sender)
  updateTextInput(session, inputId = 'comment', label = 'Comment', value = '')
  updateTextInput(session, inputId = 'subject', label = 'Subject',
                  value = global$subject)
  updateTextInput(session, inputId = 'location', label = 'Location', value = '',
                  placeholder = 'gridref of observation')
  updateTextInput(session, inputId = 'tel', label = 'Telephone Number',
                  value = '')
  updateTextInput(session, inputId = 'subject_reply', label = 'Subject',
                  value = paste('Re:',global$subject))
  updateTextAreaInput(session, inputId = 'correspondance',
                      label = 'Correspondence', value = global$msgbody)
  updateTextInput(session, inputId = 'location_description',
                  label = 'Location Description', value = '')
  updateSelectInput(session, inputId = 'expert', label = 'Expert Knowledge?',
                    choices = c('',
                                'General nature recording',
                                'Entomology',
                                'Apiculture'),
                    selected = '')
  updateTextInput(session, inputId = 'date', label = 'Date',
                  value = as.character(global$date))
  updateTextInput(session, inputId = 'i', label = 'Select Index (i)',
                  value = as.character(values$i))
  updateSelectInput(session, inputId = 'email_text_selector',
                    label = 'Email Response',
                    choices = c('Custom','Giant Woodwasp',
                                'Hoverfly','European Hornet'),
                    selected = 'Custom')
  updateTextAreaInput(session, inputId = 'email_text',
                      label = 'Message Body', value = '')
  output$geotable <- renderDataTable({
    NULL
  })
  output$sendemail <- renderText({''})
  output$serverResponse <- renderText({''})
  list(output = return_list$output,
       values = return_list$values,
       global = global)
}
# Get the contents of the email

# Function to send an email
# Function requires:
#   OutApp - an outlook object
#   values - shiny reactive values list
#   reply - the main body reply for the message
# Function returns:
#   values
# This function currently does not change the values list, but it is passed back to
# caller in case changes are required in the future
send_email <- function(OutApp, values, reply, recipient, msgBody, subject,
                       from = NULL){
  # create an email 
  outMail = OutApp$CreateItem(0)
  outMail[["To"]] = recipient
  outMail[["subject"]] = subject
  outMail[["body"]] = msgBody
  if(!is.null(from)){
    outMail[["SentOnBehalfOfName"]] <- from
  }
  ### send it
  outMail$Send()
  values
}

# Function to take an outlook Name Space and return the full list of folders
find_folders <- function(outlookNameSpace){
  numNames <- outlookNameSpace$Folders()$Count()
  foldersDF <- lapply(1:numNames, FUN = function(i){
    userName <- outlookNameSpace$Folders(i)$Name()
    numFolders <- outlookNameSpace$Folders(i)$Folders()$Count()
    if(numFolders > 0){
      folders <- lapply(1:numFolders, FUN = function(j){
        folderName <- outlookNameSpace$Folders(i)$Folders(j)$Name()
        numSubFolders <- outlookNameSpace$Folders(i)$Folders(j)$Folders()$Count()
        if(numSubFolders > 0){
          folderDF <- lapply(1:numSubFolders, FUN = function(k){
            subFolderName <- outlookNameSpace$Folders(i)$Folders(j)$Folders(k)$Name()
            numSubSub <- outlookNameSpace$Folders(i)$Folders(j)$Folders(k)$Folders()$Count()
            data.frame(i = i, j = j, k = k,
                       numSubSub = numSubSub,
                       subFolderName = subFolderName)
          }) %>% bind_rows()
        } else {
          folderDF <- data.frame(i = i, j = j, k = NA,
                                 numSubSub = NA,
                                 subFolderName = NA)
        }
        folderDF$folderName <- folderName
        folderDF
      }) %>% bind_rows()
    } else {
      folders <- data.frame(i = i, j = NA, k = NA,
                            numSubSub = NA,
                            subFolderName = NA, folderName = NA)
    }
    folders$userName <- userName
    folders
  }) %>% bind_rows()
  foldersDF
}

getpostcode <- function(txt){
  postcode <- str_extract_all(txt, paste0(
    '([Gg][Ii][Rr] 0[Aa]{2})|',
    '((([A-Za-z][0-9]{1,2})|',
    '(([A-Za-z][A-Ha-hJ-Yj-y][0-9]{1,2})|',
    '(([A-Za-z][0-9][A-Za-z])|',
    '([A-Za-z][A-Ha-hJ-Yj-y][0-9][A-Za-z]?)))) ?[0-9][A-Za-z]{2})')) %>%
    unlist()
  if(length(postcode)>0){
    return(data.frame(lat = NA, lng = NA, name = postcode))
  } else {
    return(NULL)
  }
}

stopwrds <- c(stopwords('en'), 'would')
geoparse <- function(textl, l){
  output <- tryCatch({
    op <- geonames::GNsearch(name = textl[l])
    op <- op[op$countryName == 'United Kingdom',]
    if(nrow(op)>0){
      op <- op %>% select('lat','lng','toponymName')
      tf <- lapply(op$toponymName, FUN = function(namep){
        tnl <- strsplit(namep,split = ' ') %>% unlist()
        tnl[1] == textl[l]
      }) %>% unlist()
      op <- op[tf,]
      return(op)
    } else {
      return(NULL)
    }
  }, error = function(e){
    return(NULL)
  })
}

getPage <- function(URLgeo) {
  return((browseURL(URLgeo)))
}

getPage2 <- function(URLgeo) {
  return((HTML(readLines(URLgeo))))
}

getPage4 <- function(URLgeo) {
  return(tags$script(HTML(httr::content(GET(URLgeo), 'text'))))
}

myPopify <- function(bs, txt){
  popify(el = bs, title =  '', placement = 'bottom', content = txt,
         trigger = 'hover',  options = list(container = 'body'))
}

# Some code to search - saving for later development
#search.phrase <- '2020-07-09'
#search <- OutApp$AdvancedSearch(
#  "Inbox",
#  paste0("http://schemas.microsoft.com/mapi/proptag/0x0037001E ci_phrasematch '", search.phrase, "'")
#)
#results <- search[['Results']]
#results[[1]][['Subject']]
#results$Count()
