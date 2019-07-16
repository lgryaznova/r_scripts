# Glossary builder for a smartCAT translation project (1C SSL)

# Use this script to extract namespace and its translation
# from completed smartCAT .docx bilingual files saved in a directory. 
# Save all files to a directory, put the script to the same directory 
# for your convenience, and run it.

# Returns a glossary in a .csv file for each .docx file. The .csv
# file contains a table of two columns ("Russian" and "English")
# with the code namespace and its translation for the related
# .docx file.

# DO NOT use with uncompleted translation as there is no check for
# empty strings or Cyrillic symbols in the translation column yet.
# The script is intended for processing completed files only.


# set Russian locale
Sys.setlocale(,"ru_RU")

# install packages if missing
for (package in c('docxtractr', 'dplyr')) {
    if (!require(package, character.only=T, quietly=T)) {
        install.packages(package)
        library(package, character.only=T)
    }
}

# process .docx files from the current directory
filenames <- Sys.glob("*.docx")
for (fn in filenames) {
    
    # read all tables from docx
    doc <- read_docx(fn)
    tables <- data.frame(docx_extract_all_tbls(doc))
    names(tables) <- tables[1,]
    
    # remove unnecessary columns and a colNames row, rename columns
    tables <- tables[2:3]
    tables <- tables[2:dim(tables)[1],]
    names(tables) <- c('Russian', 'English')
    
    # filter segments with no spaces and punctuation in Russian
    tables$nospace <- grepl('^[А-Яа-я]*$', tables$Russian)
    tables <- filter(tables, nospace == T)
    
    # order alphabetically, leave unique rows only, remove 'nospace' column
    tables <- tables[order(tables$Russian),]
    tables <- unique(tables)
    tables <- tables[1:2]
    
    # write glossaries to csv, add '_glossary.csv' to original filenames
    write.csv(tables, file = paste(fn, '_glossary.csv', sep = ''),
              row.names = F)
}
