library(httr) # install.packages("httr")

username <- "username"
password <- "password"

query <- "https://api.connect.ihsmarkit.com/dataplatform/v1/odata/ChemicalCapacityByCompany_CapacityToProduce_API_Data_Simplified?$top=3000"

while (!is.null(query) && query!= "") {
  response <- GET(query, authenticate(username, password, type = "basic"))
  stop_for_status(response)
  warn_for_status(response)
  
  response_content <- content(response, encoding = "UTF-8")
  print(response_content)
  
  query <- toString(response_content[["@odata.nextLink"]])
}