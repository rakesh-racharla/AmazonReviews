#' @title  pulls product reviews from Amazon.com.
#'
#' @description pull reviews/opinions from Amazon and store in excel sheet.
#'
#' @param url url of the products second page
#' @param n total number of pages.
#' @param filename name of the excel file.
#' @param sheetname name of the sheet of excel file.
#'
#'@depends XML,httr,xlsx,stringr,utils
#'
#' @return Dataframe
#'
#' @examples extract.amazon.reviews("https://www.amazon.com/Rich-Dad-Poor-Teach-Middle/product-reviews/1612680194/ref=cm_cr_arp_d_paging_btm_2?ie=UTF8&reviewerType=all_reviews&pageNumber=2",3,"reviews","rich_dad_poor_dad")
#'
#' @export extract.amazon.reviews
#'
#'

extract.amazon.reviews= function(url,    # 2nd page review url of product
                    n,filename,sheetname)      # Number of pages to extarct

{
  req.packages <- c("XML","httr","xlsx","stringr","utils")
  new.pkg <- req.packages[!(req.packages %in% installed.packages()[, "Package"])]
  if (length(new.pkg))
    install.packages(new.pkg, dependencies = TRUE)
  sapply(req.packages, require, character.only = TRUE)

  text_page=character(0)  # Create blank document

  pb <- txtProgressBar(min = 1, max = n, style = 3)    # define progress bar

  url = unlist(str_split(url,"ie="))[1]   # Process url
  url = substr(url,1,nchar(url)-2)        # Process url

  for(p in 1:n){

    url0 = paste(url,p,"?ie=UTF8&pageNumber=",p,"&showViewpoints=0&sortBy=byRankDescending",sep="") # Create final url

    textr = htmlParse(rawToChar(GET(url0)$content))   # read Url
    result <- lapply(textr['//span[@class="a-size-base review-text"]'],xmlValue)
    text_page = c(text_page,result)


    setTxtProgressBar(pb, p)             # print progress bar

    #Sys.sleep(1)
  }

  text_page =gsub("<.*?>", "", text_page)       # regex for Removing HTML character

  text_page = gsub("^\\s+|\\s+$", "", text_page) # regex for removing leading and trailing white space
  res <- unlist(text_page)
  res <- as.data.frame(res)
  write.xlsx(res,file=paste(filename,".xlsx",sep = ""),sheetName=sheetname)
  print(paste("file saved as ",filename,".xlsx in ",getwd(),sep = ""))
}
