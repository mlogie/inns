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
  values$subject <- global$subject <- emails(datesDF$j[values$i])[['Subject']]
  values$msgbody <- global$msgbody <- emails(datesDF$j[values$i])[['Body']]
  values$date <- global$date <- getDate(emails(datesDF$j[values$i]))
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
    if(grepl('jpg$',values$attachments)){
      output$myImage <- renderImage({
        createJPEG(attachment_file)
      }, deleteFile = TRUE)
    } else if(grepl('png$',values$attachments)){
      output$myImage <- renderImage({
        createPNG(attachment_file)
      }, deleteFile = TRUE)
    } else {
      output$myImage <- renderImage({
        createPNG('www/unknown_format.png', height = 100, width = 250)
      }, deleteFile = FALSE)
      dispName <- attach_obj$Item(values$img_num)[['DisplayName']]
      newName <- file.path(dirname(attachment_file),
                           dispName)
      file.rename(attachment_file,newName)
      output$attachment_info <- renderText({
        paste0('Unknown format for attachment: ',
               values$attachments,', ',
               values$img_num,' of ',values$num_attachments,
               '.\nFile saved here:',newName)
      })
    }
  } else {
    values$num_attachments <- 0
    values$attachments <- ''
    output$myImage <- renderImage({
      createPNG('www/no_attachments.png', height = 100, width = 250)
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
  updateTextInput(session, inputId = 'subject', label = 'Subject',
                  value = global$subject)
  updateTextInput(session, inputId = 'date', label = 'Date',
                  value = as.character(global$date))
  updateTextInput(session, inputId = 'i', label = 'Select Index (i)',
                  value = as.character(values$i))
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
send_email <- function(OutApp, values, reply, recipient){
  # create an email 
  outMail = OutApp$CreateItem(0)
  outMail[["To"]] = recipient
  outMail[["subject"]] = paste0('Re:',values$subject)
  outMail[["body"]] = paste0('Thank you for your email ',values$sender,reply)
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
