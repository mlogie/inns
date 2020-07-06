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
  plot(1,1,xlim=c(1,res[1]),ylim=c(1,res[2]),asp=1,
       type='n',xaxs='i',yaxs='i',xaxt='n',yaxt='n',xlab='',ylab='',bty='n')
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

# Function to grab the 
# Function requires:
#   emails - RDCOM pointer to list of emails
#   values - shiny reactive values list
# Function returns:
#   values
# This function updates:
#   values$subject - subject
#   values$sender - sender
#   values$msgbody - message body
# The email for which information extracted is the value saved as values$i
extract_contents <- function(emails, values, global){
  values$sender <- global$sender <- tryCatch({
    if(emails(values$i)$Sender()$AddressEntryUserType() == 30){
      emails(values$i)[['SenderEmailAddress']]
    } else {
      emails(values$i)[['Sender']][['GetExchangeUser']][['PrimarySmtpAddress']]
    }
  }, error = function(error) {
    return('Could not obtain sender')
  })
  values$subject <- global$subject <- emails(values$i)[['Subject']]
  values$msgbody <- global$msgbody <- emails(values$i)[['Body']]
  values$date <- global$date <- as.Date("1899-12-30") + floor(emails(values$i)$ReceivedTime())
  list(values = values, global = global)
}

# Function which grabs an attachment from an email, saving the JPEG/PNG attachment
# to output$myImage as a renderImage shiny object
# Function requires:
#   emails - RDCOM pointer to list of emails
#   values - shiny reactive values list
#   output - shiny output list
# Function returns a list containing: (output, values)
# This function updates:
#   values$attachments - the name of the current attachment (based on values$img_num)
#   values$num_attachments - the number of attachments
#   output$attachment_info - shiny renderText with info about the attachment
#     Options are:
#       "Attachment x of y"
#       "Unknown image format for attachment"
#       "No attachments"
format_attachments <- function(emails, values, output){
  attach_obj <- emails(values$i)[['attachments']]
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
      
      output$attachment_info <- renderText({
        paste0('Unknown image format for attachment: ',
               values$attachments,', ',
               values$img_num,' of ',values$num_attachments)
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

# Function to send an email
# Function requires:
#   OutApp - an outlook object
#   values - shiny reactive values list
#   reply - the main body reply for the message
# Function returns:
#   values
# This function currently does not change the values list, but it is passed back to
# caller in case changes are required in the future
send_email <- function(OutApp, values, reply){
  # create an email 
  outMail = OutApp$CreateItem(0)
  outMail[["To"]] = "marlog@ceh.ac.uk"
  outMail[["subject"]] = paste0('Re:',values$subject)
  outMail[["body"]] = paste0('Thank you for your email ',values$sender,reply)
  ### send it                     
  outMail$Send()
  values
}