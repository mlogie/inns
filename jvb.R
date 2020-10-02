# Function to get security nonce and authentication token
# Takes:
#   URLnonce: set to the dev warehouse url for now
#   password: user password
# Returns: text string to append to website for posting
getnonce <- function(URLnonce = 'http://devwarehouse.indicia.org.uk/index.php/services/security/get_nonce',
                     password){
  r <- POST(URLnonce,
            body = list(website_id = 109))
  nonce <- httr::content(x = r, as = 'text')
  key <- paste0(nonce, ':', password)
  authtoken <- digest(key, 'sha1', serialize = FALSE)
  
  URLappend <- paste0('?auth_token=', authtoken,
                      '&nonce=', nonce,
                      '&website_id=109')
  
  return(URLappend)
}

# Function to post a json to the data warehouse
# Takes:
#   URLauth: the URL string from function getnonce()
#   submission: the sample, in json format
# Returns: the content of the return message from the warehouse
postsubmission <- function(URLauth, submission){
  URL <- paste0('http://devwarehouse.indicia.org.uk/index.php/services/data/save',
                URLauth)
  
  r <- httr::POST(URL,
                  body = list('submission' = I(submission)))
  return(httr::content(x = r, as = 'text'))
}

# Function to post an image to the data warehouse
# Takes:
#   URLauth: the URL string from function getnonce()
#   imgpath: the path to the image
# Returns: the image path from the server e.g. '123456789image.png'
postimage <- function(URLauth, imgpath){
  URLimg <- paste0('http://devwarehouse.indicia.org.uk/index.php/services/data/handle_media',
                   URLauth)
  
  res <- POST(url=URLimg,
              body=list('media_upload'=upload_file(imgpath)))
  return(httr::content(x = res, as = 'text'))
}

# Function to take some parameters and turn it into a valid json format
# Takes:
#   imgString: the image string returned from function postimage. This can
#              be a character class of a vector or list of image strings.
#              Each one passed will be added to the occurrence.
#   email: email address of the source
#   tel: telephone number of the source
#   correspondance: text with information from source (I know I spelled this
#                   incorrectly, but it's how it's spelled in the warehouse)
# Returns: nicely formatted json
createjson <- function(imgString = NULL, email = NULL,
                       tel = NULL, date = NULL, location = '',
                       experience = 14684, correspondance = '',
                       comment = '', location_description = ''){
  if(!grepl(pattern = '^[A-Z]+[0-9]+$',location)){
    return('Location improperly formatted')
  }
  #  "smpAttr:1304":{"value":"Enter recorder experience"}
  experience <- switch(experience,
                       'General nature recording' = 14684,
                       'Entomology' = 14685,
                       'Apiculture' = 14686)

  # Create the sample fields
  fields <- list(website_id = list(value = "109"),
                 survey_id = list(value = "500"),
                 entered_sref = list(value = location),
                 entered_sref_system = list(value = "OSGB"),
                 location_name = list(value = location_description),
                 comment = list(value = comment),
                 `smpAttr:1140` = list(value = ""),
                 `smpAttr:43` = list(value = "TRUE"))
  if(!is.null(tel)){
    fields$`smpAttr:20` = list(value = tel)
  }
  if(!is.null(email)){
    fields$`smpAttr:35` = list(value = email)
  }
  if(!is.null(date)){
    fields$date = list(value = date)
  }
  if(!is.null(correspondance)){
    fields$`smpAttr:1141` = list(value = correspondance)
  }
  if(!is.null(experience)){
    fields$`smpAttr:1304` = list(value = experience)
  }

  # Create the occurrence fields
  occ_fields <- list(zero_abundance = list(value = "f"),
                     taxa_taxon_list_id = list(value = "289248"),
                     website_id = list(value = "109"),
                     record_status = list(value = "C"),
                     confidential = list(value = "t"),
                     sensitivity_precision = list(value = "100km"))
  occurrence <- list(list(fkId = "sample_id",
                          model = list(id = "occurrence",
                                       fields = occ_fields)))
  
  # For every image supplied, create an image instance
  if(!is.null(imgString)){
    media <- lapply(imgString, FUN = function(imgx){
      med_fields <- list(path = list(value = imgx),
                         caption = list(value = "Enter comment here"))
      list(fkId = "occurrence_id",
           model = list(id = "occurrence_medium",
                        fields = med_fields))
    })
    # Add the images to the occurrence as a submodel
    occurrence[[1]]$model$subModels <- media
  }
  
  outjson <- list(id = "sample", fields = fields, subModels = occurrence) %>%
    toJSON(auto_unbox = TRUE)
  
  return(outjson)
}
