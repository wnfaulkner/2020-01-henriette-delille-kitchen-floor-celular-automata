#0000000000000000000000000000000000000000000000000000000#
#      	2019-10 GI SNA                             	    #
#0000000000000000000000000000000000000000000000000000000#

# 0-SETUP -----------------------------------------------------------
  
  #INITIAL SETUP
    rm(list=ls()) #Remove lists
    options(java.parameters = "- Xmx20g") #helps r not to fail when importing large xlsx files with xlsx package
    
    
    #Section & Code Clocking
      sections.all.starttime <- Sys.time()
  
  # ESTABLISH BASE DIRECTORIES
  
    # Figure out what machine code is running on
      if(dir.exists("C:\\Users\\WNF\\Meu Drive")){
        m900 <- FALSE
        base.dir <- "C:\\Users\\WNF"
      }else{
        m900 <- TRUE
        base.dir <- "C:\\Users\\WNF"
      }
      
    # Set Working Directory and R Project Directory
      if(m900){  
        #M900
          wd <- paste(base.dir, "\\Google Drive\\1. FLUX PROJECTS - CURRENT\\2019-08 EXT Global Integrity SNA\\3. Data Analysis\\", sep = "")
          rproj.dir <- paste(base.dir, "\\Documents\\GIT PROJECTS\\2019-10-GI-SNA-Analysis\\", sep = "")
      }else{
        #Thinkpad T470
          wd <- "X:\\Google Drive File Stream\\Meu Drive\\1. FLUX PROJECTS - CURRENT\\2019-08 EXT Global Integrity SNA\\3. Data Analysis\\"
          rproj.dir <- "C:\\Users\\WNF\\Documents\\GIT PROJECTS\\2019-10-GI-SNA-Analysis\\"
      }
    
    #Check Directories
      wd <- if(!dir.exists(wd)){choose.dir()}else{wd}
      rproj.dir <- if(!dir.exists(rproj.dir)){choose.dir()}else{rproj.dir}
    
    #Source Tables Directory (raw data, configs, etc.)
      source.tables.dir <- paste(wd,"0. Inputs\\", sep = "")
      if(dir.exists(source.tables.dir)){ 
        print("source.tables.dir exists.")
      }else{
        print("source.tables.dir does not exist yet.")
      }
      print(source.tables.dir)
      
    #Establish Outputs Directory
      outputs.parent.folder <- 
        paste(
          wd,
          "1. Outputs\\",
          sep  = ""
        )
      
      setwd(outputs.parent.folder)
    
  # LOAD UTILS FUNCTIONS
      
    setwd(rproj.dir)
    source("99-Utils GI SNA.r")
    
  # LOAD LIBRARIES/PACKAGES
    
    #In case working on new R install that does not have packages installed
    #InstallCommonPackages()
    #install.packages("ReporteRs")
    #install.packages("jsonline")
    #install.packages('httpuv')
    #install.packages('xtable')
    #install.packages('sourcetools')
    #install.packages('shiny')
    #install.packages('miniUI')
    
    LoadCommonPackages()
    
    #Section Clocking
      #section0.duration <- Sys.time() - section0.starttime
      #section0.duration

# 1-IMPORT / CREATE TEST DATA -----------------------------------------
  
  #Section Clocking
    tic("1-Import") 
    
  #Source Import Functions
    
    #source("1-import_functions.r")
  
  #IMPORT CONFIG TABLES ----
    configs.ss <- gs_key("1_WEh68ccX0k63Szt-lmrI6sKfwXcp9JsmJJDA3mNPcU",verbose = TRUE) 
    
    #Import all tables from config google sheet as tibbles
      all.configs.ls <- GoogleSheetLoadAllWorksheets(configs.ss)
    
    #Assign each table to its own tibble object
      ListToTibbleObjects(all.configs.ls) #Converts list elements to separate tibble objects names with
                                          #their respective sheet names with ".tb" appended
    
    #Extract global configs from tibble as their own character objects
      TibbleToCharObjects(
        tibble = config.global.tb,
        object.names.colname = "config.name",
        object.values.colname = "config.value"
      )
      use.test.data <-
        ifelse(use.test.data == "true", TRUE, FALSE)

  #IMPORT RESPONSES TABLE (IF NOT GENERATING TEST DATA) ----
    if(!use.test.data){
      setwd(source.tables.dir)
      
      resp1.tb <- 
        read.csv(
          file =  
            MostRecentlyModifiedFilename(
              title.string.match = main.data.file.name.character.string,
              file.type = "csv",
              dir = source.tables.dir
            ),
          stringsAsFactors = FALSE,
          header = TRUE
        ) %>% 
        as_tibble(.)
    }
  
  #CREATE TEST DATA ----
    if(use.test.data){
      
      #num.enumerators <- 10
      #expected.num.responses <- 50
      resp.ls <- list()
      
      #i = 21
      for(i in 1:nrow(variables.tb)){ #START OF LOOP 'i' BY VARIABLE
      
        var.id.i <- variables.tb$var.id[i]
        var.type.i <- variables.tb$var.type[variables.tb$var.id == var.id.i]
        
        #if variable type not one of allowable types, skip to next loop
          allowable.var.types <- c("categorical", "binary")#, "date", "open.text", "ordinal") 
          if(!var.type.i %in% allowable.var.types){
            print(
              paste(
                "Loop: [",
                i,
                "]: Variable type in variables table not one of allowable types.",
                sep = ""
              )
            )
            next()
          }
        
        #Binary
          #when categories not specified in configs
            #if(var.type.i == "binary" && is.na(variables.tb$ans.opt[i])){
            #  var.i <- 
            #    sample(
            #      c(0,1), 
            #      replace = TRUE, 
            #      size = expected.num.responses, 
            #      prob = runif(2, min = 0, max = 1)
            #    )
            #}
          
          #when categories specified in configs        
            #if(var.type.i == "binary" && !is.na(variables.tb$ans.opt[i])){
            #  
            #  ans.opt.i <- variables.tb$ans.opt[i] %>% strsplit(., split = ";") %>% unlist %>% trimws
            #  
            #  var.i <- 
            #    sample(
            #      ans.opt.i, 
            #      replace = TRUE, 
            #      size = expected.num.responses, 
            #      prob = runif(2, min = 0, max = 1)
            #    )
            #}
          
        #Ordinal
          
        
        #Categorical
          #when categories not specified in configs
            #if(var.type.i == "categorical" && is.na(variables.tb$ans.opt[i])){
            #    
            #    var.i <-  
            #      sample(
            #        seq(from = 1, to = num.enumerators),
            #        replace = TRUE,
            #        size = expected.num.responses
            #      ) 
            
            #}
          
          #when categories specified in configs        
            #if(var.type.i == "categorical" && !is.na(variables.tb$ans.opt[i])){
            #  
            #  ans.opt.i <- variables.tb$ans.opt[i] %>% strsplit(., split = ";") %>% unlist %>% trimws
            #  
            #  #! make so can have non-uniform probabilities for different answer options 
            #  #prob.i <- 
            #  
            #  var.i <- 
            #    sample(
            #      ans.opt.i,
            #      replace = TRUE,
            #      size = expected.num.responses
            #      #prob = prob.i
            #    )
            #}
        
        #Date
          
        #Open Text
        
        resp.ls[[i]] <- var.i
        
      } #END OF LOOP 'i' BY VARIABLE
      
      #Compile list into tibble
      
      #Update names to correct variable names
      
    } #End of if-statement for creating test data
    
  #Section Clocking
    #section1.duration <- Sys.time() - section1.starttime
    #section1.duration
    #Sys.time() - sections.all.starttime
    #toc()
    
    
# 2-CLEANING -----------------------------------------
  
  #Section Clocking
    tic("2-CLEANING")
    
  #Source Cleaning Functions
    cleaning.source.dir <- paste(rproj.dir,"2-Cleaning/", sep = "")
    setwd(rproj.dir)
    
  #CLEANING ANSWER OPTIONS TABLE ----
    ans.opt.tb <- 
      apply(
        ans.opt.tb,
        2,
        function(x){SubNA(vector = x, na.replacement = "")}
      ) %>%
      as_tibble() %>%
      mutate(
        ans.opt.set.id = as.numeric(ans.opt.set.id),
        ans.opt.numeric = as.numeric(ans.opt.clean_num)
      )
    
  #CLEANING RESPONSE TABLE ----  
    
    #Names/Column Headers
      
      #Replace imported variable names with var.ids (requires some munging of character strings to make them match)
        resp2.tb <- 
          LowerCaseNames(resp1.tb) %>%  #Lower-case all variable names
          as_tibble()
          
        names(resp2.tb) <- SubRepeatedCharWithSingleChar(string.vector = names(resp2.tb), char = ".")
        
        names(resp2.tb)[names(resp2.tb) %>% equals("enumerator.s.notes.1") %>% which] <- 
          c("enumerator.s.notes.1.asset.tracing", "enumerator.s.notes.1.data.sources.challenges")
        
        names(resp2.tb) <- 
          IndexMatchToVectorFromTibble(
            vector = names(resp2.tb),
            lookup.tb = variables.tb,
            match.colname = "raw.var.imported",
            replacement.vals.colname = "var.id",
            mult.replacements.per.cell = FALSE,
            print.matches = FALSE
          )
        
    #Add 'id' variable
      resp2.tb %<>%
        mutate(
          resp.id = 1:nrow(.)
        )
    
    #Remove Blank Variables
      resp3.tb <-
        resp2.tb[
          ,
          apply(resp2.tb, 2, function(x){x %>% unique %>% is.na %>% all}) %>% unlist %>% as.vector %>% not
        ]
      
    #Lower-case all character variable data
      resp4.tb <- LowerCaseCharVars(resp3.tb)  
      
    #Fix Typos in data entry
      
      #'inbetween' > 'in between'
        resp5.tb <-
          TableGsub(
            resp4.tb, 
            "inbetween not part and somewhat part of my responsibilities",
            "in between not part and somewhat part of my responsibilities"
          ) %>% 
          TableGsub(
            ., 
            "in between part and somewhat of my responsibilities",
            "in between not part and somewhat part of my responsibilities"
          ) %>%
          as_tibble()
        
    #Convert Time/Date Variables and create separate Time and Date variables
      resp6.tb <-
        resp5.tb %>%
        mutate(
          data.entry.date =
            resp5.tb$timestamp %>%
            strsplit(., " ") %>%
            lapply(., '[[', 1) %>%
            unlist %>%
            as.Date(., format = "%m/%d/%Y" ),
          data.entry.time = 
            resp5.tb$timestamp %>%
            strsplit(., " ") %>%
            lapply(., '[[', 2) %>%
            unlist %>%
            strptime(., format = "%H:%M:%S") %>%
            format(., "%H:%M:%S"),
          interview.date =
            resp5.tb$interview.date %>%
            strptime(., format = "%m/%d/%Y") %>%
            format(., "%Y-%m-%d")
        )
      
    #Specific Variable Recoding
      
      #Interview Location
        
        #create online binary variable
          resp6.tb %<>%
            mutate(
              interview.location.online = grepl("online", resp6.tb$interview.location)
            )
        
        #standardize port harcourt entries
          resp6.tb$interview.location[grep("harcourt", resp6.tb$interview.location)] <- 
            "port harcourt"
          
      #Work Location (State)
        resp6.tb$work.location.state %<>%
          gsub(
            "niger delta region",
            "abia,akwa ibom,bayelsa,cross river,delta,edo,imo,ondo,rivers", 
            .
          )
        
        resp6.tb$reported.case.authority.level %<>%
          gsub(
            "states / subnational",
            "states/subnational",
            .
          )
        
        resp6.tb$reported.case.authority.level %<>%
          gsub(
            "states / subnational",
            "states/subnational",
            .
          )
        
    #Create Equivalent Numeric Variables for binary, some ordinal, and some scalar variables
      create.numeric.varnames <-
        variables.tb %>% 
        filter(create.numeric.var) %>%
        select(var.id) %>%
        unlist %>% as.vector
          
      create.numeric.tb <- 
        resp6.tb %>% 
        select(c("resp.id", create.numeric.varnames))
      
      
      resp.long.ls <- list()
      resp.other.ls <- list()
      
      #i=2 #for loop testing
      for(i in 1:ncol(create.numeric.tb)){
        
        #Loop Inputs
          varname.i <- names(create.numeric.tb)[i]  
          var.i <- create.numeric.tb %>% select(varname.i)
          var.type.i <- variables.tb %>% filter(var.id == varname.i) %>% select(var.type) %>% unlist %>% as.vector
          
          ans.set.id.i <-
            variables.tb %>%
            filter(var.id == varname.i) %>%
            select(ans.opt.set.id) %>%
            unlist %>% as.vector
          
          if(is.na(ans.set.id.i) || length(ans.set.id.i) == 0){
            print(i)
            #print("Answer option set id not matching between 'variables' and 'ans.opt' config tables. Skipping to next loop.")
            next()
          }else{
            print(c(i, "stored in list"))
          }
        
          ans.opt.tb.i <- 
            ans.opt.tb %>%
            filter(ans.opt.set.id == ans.set.id.i)
          
          ans.opt.raw.i <- 
            ans.opt.tb.i$ans.opt.raw #%>%
            #gsub(",","", .)
        
        #Creating long response table
          
          #All vars except mult choice mult answer
            if(var.type.i != "categorical.mult"){
              result.resp.tb <-
                tibble(
                  resp.id = create.numeric.tb$resp.id,
                  ans.set.id = ans.set.id.i,
                  var = varname.i,
                  ans.opt.raw = var.i %>% unlist,
                  value = var.i %>% unlist
                )
            }
          
          #Mult choice mult answer vars
            if(var.type.i == "categorical.mult"){
              
              #Main Responses
                result.resp.tb <- 
                  sapply(
                    ans.opt.tb.i$ans.opt.raw,
                    function(x){grepl(x, unlist(var.i))}
                  ) %>% 
                  as_tibble() %>%
                  mutate(
                    resp.id = create.numeric.tb$resp.id,
                    ans.set.id = ans.set.id.i,
                    ans.raw = var.i %>% unlist
                  )
              
              #Create separate 'other' data table for when there are 'other' responses concatenated within response strings
                resp.other.index <- length(resp.other.ls) + 1
                
                resp.other.ls[[resp.other.index]] <- 
                  gsub(
                    paste0(
                      ans.opt.tb.i$ans.opt.raw, collapse = "|"
                    ), 
                    "", 
                    unlist(var.i)
                  ) %>% 
                  as.vector() %>% 
                  gsub("\\,","",.) %>% 
                  SubRepeatedCharWithSingleChar(., " ") %>% 
                  trimws %>%
                  as_tibble %>%
                  mutate(
                    resp.id = create.numeric.tb$resp.id,
                    ans.raw = var.i %>% unlist
                  ) %>%
                  filter(value != "")
            }

        #Replace "" with "BLANK" as a placeholder (helps with reshaping to long data)    
          if("BLANK" %in% names(result.resp.tb)){
            result.resp.tb$BLANK <- var.i %>% equals("")
            #names(result.resp.tb) <- gsub("V1", "BLANK",names(result.resp.tb))
          }
        
        #Reshape & store in list for compilation outside of loop
          if(var.type.i != "categorical.mult"){
            resp.long.ls[[i]] <- #store resulting long data frame (with raw answer option, resp.id, and ans.set.id)
              melt(
                result.resp.tb, 
                variable.name = "ans.opt.raw",
                value.names = "value", 
                id.vars = c("resp.id","ans.set.id","ans.opt.raw")
              ) %>% 
              filter(!value %in% c(FALSE, ""))
          }
          
        
        
      } # End of loop 'i' by variable
      
      
      #Join responses with ans.opt.tb and keep only clean variables     
          #! MAKE SURE WHEN ADDING CHARACTER VARIABLE BACK IN THAT IT CHOOSES THE FIRST MATCH WHEN THERE ARE TWO 
          #IN ANS.OPT.TB - SEE ANS.OPT.ID == 11: HAVE TO TURN ALL 'NO' INTO BLANKS  
        
      #Informatics - build related tables (as per LucidChart)
    
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
  #M

        #resp4.ls <- list()
        
        #i=1
        #for(i in 1:ncol(resp3.tb)){
        #  print(names(resp3.tb)[i])
        #  var.i.import.tb <- resp3.tb[,i]
        #  var.i.class.import <- variables.tb$var.class.import[names(resp3.tb)[i] == variables.tb$var.id]
        #  var.i.class.clean <- 
        #    variables.tb$var.class.clean[names(resp3.tb)[i] == variables.tb$var.id] %>%
        #    strsplit(., ";") %>% unlist
        #  
        #  var.i.clean.tb <-
        #    SetColClass(
        #      tb = var.i.import.tb,
        #      colname = names(resp3.tb)[i],
        #      to.class = var.i.class.clean
        #    )
        #  
        #  resp4.ls[[i]] <- var.i.clean.tb
        #}
        
        #do.call(cbind, resp4.ls)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
      
  #Reshape to Long Data by response to CWIS question
    cwis.varnames <- FilterVector(grepl(paste0(domains, collapse = "|"), names(resp5.tb)), names(resp5.tb))
    resp6.tb <-
      melt( #Melt to long data frame for all cwis vars
        resp5.tb,
        id.vars = names(resp5.tb)[!names(resp5.tb) %in% cwis.varnames], 
        measure.vars = cwis.varnames
      ) %>% 
      filter(!is.na(value)) %>% #remove rows with no answer for cwis vars
      MoveColsLeft(., c("resp.id","unit.id")) %>% #Rearrange columns: resp.id and unit.id at the front
      as_tibble()
    
  #ADDING NEW USEFUL VARIABLES ----
    
    #Add proficiency dummy variable
      resp7.tb <- 
        resp6.tb %>%
        mutate(
          is.proficient = ifelse(value >= 4, 1, 0)
        )
      
    #Add variable for baseline, most recent, and next-to-most-recent years for each unit.id
      resp7.tb$year[resp7.tb$year == "baseline"] <- "0000"
      
#FUNCTION - could be made into function to designate alphabetic first/last/penultimate by group
      year.var.helper.tb <-
        resp7.tb %>%
        select(year, unit.id) %>%
        unique %>%
        mutate(
          is.baseline = ifelse(year == "0000", 1, 0),
          is.most.recent = ifelse(year == "2017-2018", 1, 0),
          is.current = ifelse(year == "2018-2019", 1, 0)
        ) %>%
        mutate(
          is.baseline.or.most.recent = ifelse(is.baseline == 1 | is.most.recent == 1, 1, 0),
          is.baseline.or.current = ifelse(is.baseline == 1 | is.current == 1, 1, 0),
          is.most.recent.or.current = ifelse(is.most.recent == 1 | is.current == 1, 1, 0)
        )
      
      #year.var.helper.ls <- list()
      
      #for(i in 1:length(unique(year.var.helper.tb$unit.id))){
      #  unit.id.i <- unique(year.var.helper.tb$unit.id)[i]
      #  year.var.helper.tb.i <- 
      #    year.var.helper.tb %>% 
      #    filter(unit.id == unit.id.i) %>%
      #    arrange(year) %>% 
      #    mutate(
      #      num.measurements = nrow(.),
      #      is.baseline = 0,
      #      is.most.recent = 0, 
      #      is.current = 0,
      #      is.most.recent.or.current = 0,
      #      is.baseline.or.current = 0
      #    )
      #  
      #  year.var.helper.tb.i$is.baseline[1] <- 1
      #  
      #  year.var.helper.tb.i$is.current[nrow(year.var.helper.tb.i)] <- 
      #    ifelse(unique(year.var.helper.tb.i$num.measurements) == 1, 0, 1)
      #  
      #  year.var.helper.tb.i$is.most.recent[nrow(year.var.helper.tb.i)-1] <- 
      #    ifelse(unique(year.var.helper.tb.i$num.measurements) == 1, 0, 1)
      #  
      #  year.var.helper.tb.i$is.most.recent.or.current[
      #     year.var.helper.tb.i$is.current == 1 | year.var.helper.tb.i$is.most.recent ==1
      #  ] <- 1
      #  year.var.helper.tb.i$is.baseline.or.current[
      #    year.var.helper.tb.i$is.current == 1 | year.var.helper.tb.i$is.baseline ==1
      #    ] <- 1
      #  
      #  year.var.helper.ls[[i]] <- year.var.helper.tb.i
      #}
        
      #year.var.helper.tb <- do.call(rbind, year.var.helper.ls)
      
      resp8.tb <- 
        left_join(
          resp7.tb,
          year.var.helper.tb,
          by = c("year", "unit.id")
        ) %>% 
        mutate(variable = as.character(variable))
   
  #FORM FINAL DATASETS - (A) SPLITCOLRESHAPED BY DOMAIN AND (B) RESTRICTED TO SAMPLE ----
       
    #Full dataset - no splitcolreshape by domain
      resp.full.nosplit.tb <- 
        resp8.tb
      
    #Full dataset - splitcolreshape by domain
      resp.full.split.tb <-
        left_join(
          x = resp8.tb,
          y = questions.tb %>% select(var.id, domain, practice),
          by = c("variable"="var.id")
        ) %>%
        SplitColReshape.ToLong(
          df = ., 
          id.varname = "resp.id" ,
          split.varname = "domain",
          split.char = ","
        ) %>% 
        as_tibble() %>%
        left_join(
          x = ., 
          y = domains.tb,
          by = c("domain" = "domain.id")
        )
    
    #Sample Datasets  
    
      #Scenario 1 - sample print
        if(sample.print){ 
          is.valid.sample <- FALSE
          while(!is.valid.sample){
            
            resp.sample.nosplit.tb <- 
              RestrictDataToSample(
                tb = resp.full.nosplit.tb,
                report.unit = "unit.id",
                sample.print = sample.print,
                sample.group.unit = "unit.id",
                sample.size = sample.size
              )
            
            resp.sample.split.tb <- 
              left_join(
                x = resp.sample.nosplit.tb,
                y = questions.tb %>% select(var.id, domain, practice),
                by = c("variable"="var.id")
              ) %>%
              SplitColReshape.ToLong(
                df = ., 
                id.varname = "resp.id" ,
                split.varname = "domain",
                split.char = ","
              ) %>% 
              as_tibble() %>%
              left_join(
                x = ., 
                y = domains.tb,
                by = c("domain" = "domain.id")
              )
            
            is.valid.sample <- 
              ifelse(
                dcast(
                  data = resp.sample.nosplit.tb, 
                  formula = unit.id ~ .,
                  fun.aggregate = function(x){length(unique(x))},
                  value.var = "building.id"
                ) %>% 
                select(".") %>%
                unlist %>% as.vector %>%
                is_weakly_less_than(as.numeric(max.building.count)) %>% 
                all,
                TRUE,
                FALSE
              )
          }
          
          unit.ids.sample <-
            resp.sample.nosplit.tb %>%
            select(unit.id) %>%
            unique %>%
            unlist %>% as.vector
          
        }
        
      #Scenario 2 - full print (fresh, not adding to previous full print)  
        if(!sample.print & !add.to.last.full.print){
            
          is.valid.sample <- FALSE
          while(!is.valid.sample){
            
            resp.sample.nosplit.tb <- 
              RestrictDataToSample(
                tb = resp.full.nosplit.tb,
                report.unit = "unit.id",
                sample.print = TRUE,
                sample.group.unit = "unit.id",
                sample.size = sample.size
              )
            
            resp.sample.split.tb <- 
              left_join(
                x = resp.sample.nosplit.tb,
                y = questions.tb %>% select(var.id, domain, practice),
                by = c("variable"="var.id")
              ) %>%
              SplitColReshape.ToLong(
                df = ., 
                id.varname = "resp.id" ,
                split.varname = "domain",
                split.char = ","
              ) %>% 
              as_tibble() %>%
              left_join(
                x = ., 
                y = domains.tb,
                by = c("domain" = "domain.id")
              )
            
            unit.ids.sample <-
              resp.sample.nosplit.tb %>%
              select(unit.id) %>%
              unique %>%
              unlist %>% as.vector
            
            is.valid.sample <- 
              ifelse(
                length(unit.ids.sample) == as.numeric(sample.size),
                TRUE,
                FALSE
              )
            
            print(is.valid.sample)
          }
          
        }
        
      #Scenario 3 - full print (adding to previous full print)  
        if(!sample.print & add.to.last.full.print){
          is.valid.sample <- FALSE
          while(!is.valid.sample){
            
            unit.ids.sample <-
              resp.full.nosplit.tb %>%
              select(unit.id) %>%
              unique %>%
              unlist %>% as.vector %>%
              setdiff(., districts.that.already.have.reports) %>%
              .[1:sample.size] %>%
              RemoveNA
            
            #if(length(districts.without.reports) > sample.size){unit.ids.sample <- districts.without.reports[1:sample.size]}
            #if(length(districts.without.reports) < sample.size){unit.ids.sample <- districts.without.reports}
            
            is.valid.sample <-
              unit.ids.sample %in% districts.that.already.have.reports %>% not %>% all
            
            print(is.valid.sample)
            if(!is.valid.sample){next()}
            
            resp.sample.nosplit.tb <- 
              resp.full.nosplit.tb %>%
              filter(unit.id %in% unit.ids.sample)
            
            resp.sample.split.tb <- 
              left_join(
                x = resp.sample.nosplit.tb,
                y = questions.tb %>% select(var.id, domain, practice),
                by = c("variable"="var.id")
              ) %>%
              SplitColReshape.ToLong(
                df = ., 
                id.varname = "resp.id" ,
                split.varname = "domain",
                split.char = ","
              ) %>% 
              as_tibble() %>%
              left_join(
                x = ., 
                y = domains.tb,
                by = c("domain" = "domain.id")
              )
          }
        }

    #resp.long.tb <- resp10.tb
    
  #Section Clocking
    toc()
    Sys.time() - sections.all.starttime

# 3-CONFIGS (TAB, TABLE, TEXT CONFIG TABLES) ------------------
  
  #Section Clocking
    tic("3-Configs")
          
  #Load Configs Functions
    setwd(rproj.dir)
    source("3-configs_functions.r")
   
  #EXPAND CONFIG TABLES FOR EACH unit.id ACCORDING TO LOOPING VARIABLES
  
  ###                          ###
# ### LOOP "b" BY REPORT.UNIT  ###
  ###                          ###
    
  #Loop Outputs 
    config.tabs.ls <- list()
    config.tables.ls <- list()
    config.text.ls <- list()

  
  #Loop Measurement - progress bar & timing
    progress.bar.b <- txtProgressBar(min = 0, max = 100, style = 3)
    maxrow.b <- length(unit.ids.sample)
    b.loop.startime <- Sys.time()
  
  #b <- 1 #LOOP TESTER (19 = "Raytown C-2")
  #for(b in c(1,2)){   #LOOP TESTER
  for(b in 1:length(unit.ids.sample)){   #START OF LOOP BY REPORT UNIT
    
    #Create unit.id.b (for this iteration)
      unit.id.b <- unit.ids.sample[b]
    
    #Print loop messages
      if(b == 1){print("FORMING TAB, GRAPH, AND TABLE CONFIG TABLES...")}
    
    #Create data frames for this loop - restrict to unit.id id i  
      resp.long.tb.b <- 
        resp.sample.nosplit.tb%>% filter(unit.id == unit.id.b)
      
    #Other useful inputs for forming config tables
      district.name.v <- 
        resp.long.tb.b$unit.id %>% unique %>% unlist %>% as.vector %>%
        str_split(., "-") %>% lapply(., Proper) %>% unlist %>% as.vector
      
      if(length(district.name.v) > 1){
        district.name.v[length(district.name.v)] <- district.name.v[length(district.name.v)] %>% toupper()
      }
      
      district.name <- district.name.v %>% paste(., collapse = "-")
      
      tab4.loopvarname <- 
        config.tab.types.tb %>% 
        select(tab.loop.var.1) %>% 
        unlist %>% unique %>% RemoveNA
      
      building.names <- resp.long.tb.b %>% select(tab4.loopvarname) %>% unlist %>% unique
      
      building.names.tb <- 
        tibble(
          tab.type.id = 4,
          loop.id = building.names
        )
      
    #Tab config table for this report unit
      
      config.tabs.ls[[b]] <-
        tibble(
          tab.type.id = 4,
          tab.type.name = "Building Summary",
          loop.id = building.names 
        ) %>%
        rbind(
          config.tab.types.tb %>% select(tab.type.id, tab.type.name) %>% mutate(loop.id = NA) %>% filter(tab.type.id != 4),
          .
        ) %>%
        mutate(
          tab.name = 
            c(
              tab.type.name[1:3],
              paste(
                "Building Summary (",
                1:length(building.names),
                ")",
                sep = ""
              )
            )
        )
         
    #Tables config table for this report unit
      #config.table.types.tb
      config.tables.ls[[b]] <- 
        full_join(
          config.tabs.ls[[b]],
          config.table.types.tb,
          by = "tab.type.id"
        ) %>%
        mutate(
          is.state.table = !grepl("unit.id", filter.varname),
          is.domain.table = grepl("domain", filter.varname)|grepl("domain", row.header.varname)|grepl("domain",col.header.varname)
        )
        
    #Text configs table for this report unit
      
      config.text.ls[[b]] <-
        full_join(
          config.text.types.tb,
          building.names.tb,
          by = "tab.type.id"
        ) %>% 
        full_join(
          .,
          config.tabs.ls[[b]],
          by = c("tab.type.id","loop.id")
        ) %>%
        mutate(text.value = district.name) %>%
        mutate(
          text.value = 
            ifelse(text.type == "building", loop.id, text.value)
        ) %>%
        mutate(text.value = Proper(text.value))
        
    setTxtProgressBar(progress.bar.b, 100*b/maxrow.b)
    
  } # END OF LOOP 'b' BY REPORT.UNIT
  
  close(progress.bar.b)
    
  #Section Clocking
    toc()
    Sys.time() - sections.all.starttime

# 4-SOURCE DATA TABLES --------------------------------------------
  
  #Section Clocking
    tic("Section 4 duration")
        
  #Load Configs Functions
    setwd(rproj.dir)
    source("4-source data tables functions.r")
    
  #STATE AVERAGE TABLES ----
    
    #Tabs 1 & 2 ----

      #Loop Inputs
        config.tables.tab12.input.tb <- 
          config.tables.ls[[1]] %>% 
          filter(!is.na(table.type.id)) %>%
          filter(grepl("1|2", tab.type.id)) %>%
          filter(is.state.table) %>%
          OrderDfByVar(., order.by.varname = "tab.type.id", rev = FALSE) %>%
          as_tibble()
        
        tables.tab12.state.ls <- list()
      
      #d <- 4
      for(d in 1:nrow(config.tables.tab12.input.tb)){ ### START OF LOOP "d" BY TABLE ###
        
        #Print loop messages
        print(paste("TABS 1 & 2 LOOP - Loop #: ", d, " - Pct. Complete: ", 100*d/nrow(config.tables.tab12.input.tb), sep = ""))
        
        #Define table configs for loop
        config.tables.tb.d <- config.tables.tab12.input.tb[d,]
        
        #CREATE TABLE
        
          #Tab 1 Tables (newly created)
            if(config.tables.tb.d$tab.type.id == 1){
              
              #Define table aggregation formula
                table.formula.d <-
                  DefineTableRowColFormula(
                    row.header.varnames = strsplit(config.tables.tb.d$row.header.varname, ",") %>% unlist %>% as.vector,
                    col.header.varnames = strsplit(config.tables.tb.d$col.header.varname, ",") %>% unlist %>% as.vector
                  )
    
              #Define table source data
                filter.varnames.d <- config.tables.tb.d$filter.varname %>% strsplit(., ";") %>% unlist %>% as.vector
                filter.values.d <- config.tables.tb.d$filter.values %>% strsplit(., ";") %>% unlist %>% as.vector
              
                is.state.table <- ifelse(!"unit.id" %in% filter.varnames.d, TRUE, FALSE)
                is.domain.table <- grepl("domain", c(filter.varnames.d, table.formula.d)) %>% any
              
              #STATE data with NO domains in formula
                if(is.state.table & !is.domain.table){table.source.data <- resp.full.nosplit.tb}
              
              #STATE data WITH domains in formula
                if(is.state.table & is.domain.table){table.source.data <- resp.full.split.tb}
              
              #NON-STATE data NO domains in formula
                if(!is.state.table & !is.domain.table){table.source.data <- resp.sample.nosplit.tb}
              
              #NON-STATE data WITH domains in formula
                if(!is.state.table & is.domain.table){table.source.data <- resp.sample.split.tb}
              
              #Define table filtering vector
                table.filter.v <-
                  DefineTableFilterVector(
                    tb = table.source.data,
                    filter.varnames = filter.varnames.d,
                    filter.values = filter.values.d
                  )
    
              #Form final data frame
                table.d <-  
                  table.source.data %>%
                  filter(table.filter.v) %>%
                  dcast(
                    ., 
                    formula = table.formula.d, 
                    value.var = config.tables.tb.d$value.varname, 
                    fun.aggregate = table.aggregation.function
                  ) %>%
                  .[,names(.)!= "NA"]
    
              #Modifications for specific tables
                if(grepl("building.level", table.formula.d) %>% any){
                  table.d <- 
                    left_join(
                      building.level.order.v %>% as.data.frame %>% ReplaceNames(., ".", "building.level"), 
                      table.d,
                      by = "building.level"
                    )
                }
                
                if(!config.tables.tb.d$row.header){  #when don't want row labels
                  table.d <- table.d %>% select(names(table.d)[-1])
                }
            }
          
          #Tab 2 Tables (copies of tab 1)
            if(config.tables.tb.d$tab.type.id == 2){
              table.d <- 
                tables.tab12.state.ls[
                  tables.tab12.state.ls %>%
                    lapply(
                      .,
                      function(x){
                        (
                          x$configs$tab.type.id %>%
                            unlist %>% as.vector %>%
                            equals(1)
                        ) &
                          (
                            x$configs$table.type.id %>%
                              unlist %>% as.vector() %>%
                              equals(config.tables.tb.d$table.type.id)
                          )
                      }
                    ) %>%
                    unlist %>% as.vector
                  ] %>%
                .[[1]] %>% 
                .[["table"]]
            }
          
        #Table Storage
          table.d.storage.index <- length(tables.tab12.state.ls) %>% add(1)
          tables.tab12.state.ls[[table.d.storage.index]] <- list()
          tables.tab12.state.ls[[table.d.storage.index]]$configs <- config.tables.tb.d
          tables.tab12.state.ls[[table.d.storage.index]]$table <- table.d
        
      } ### END OF LOOP "d" BY TABLE ###
      
      #Finalizing loop outputs  
        names(tables.tab12.state.ls) <- 
          paste(
            #rep(unit.id.c, length(tables.tab12.state.ls)),
            #".",
            c(1:length(tables.tab12.state.ls)),
            ".",
            config.tables.tb.d$table.type.name, 
            sep = ""
          ) %>%
          gsub("-", " ", .) %>%
          gsub(" ", ".", .) %>%
          SubRepeatedCharWithSingleChar(., ".") %>%
          tolower
    
    
    #Tab 3 ----
    
      #Current vs. previous school year
        tab3.state.avg.current.vs.previous.school.year <-
          resp.full.split.tb %>% 
          filter(
            is.most.recent.or.current == 1
          ) %>%
          SplitColReshape.ToLong(
            df = ., 
            id.varname = "resp.id", 
            split.varname = "domain", 
            split.char = ","
          ) %>%
          as_tibble() %>%
          dcast(
            data = .,
            formula = domain ~ is.current,
            fun.aggregate = mean,
            value.var = "value"
          ) %>%
          ReplaceNames(., current.names = c("0","1"), new.names = c("Previous School Year","2018-2019")) %>%
          mutate(
            Trend = .[,3] - .[,2]
          ) %>%
          TransposeTable(., keep.first.colname = FALSE)
        
        tab3.state.avg.current.vs.previous.school.year %<>%
          apply(
            X = tab3.state.avg.current.vs.previous.school.year[,2:ncol(tab3.state.avg.current.vs.previous.school.year)], 
            MARGIN = 2, 
            FUN = as.numeric
          ) %>%
          cbind(
            tab3.state.avg.current.vs.previous.school.year[,1],
            .
          ) %>%
          ReplaceNames(., "Var.1", "")
        
        tab3.state.avg.current.vs.previous.school.year[1,1] <- "Previous School Year"
        
      #Current vs. baseline
        tab3.state.avg.current.vs.baseline <- 
          resp.full.split.tb %>% 
          filter(
            is.baseline.or.current == 1
          ) %>%
          SplitColReshape.ToLong(
            df = ., 
            id.varname = "resp.id", 
            split.varname = "domain", 
            split.char = ","
          ) %>%
          as_tibble() %>%
          dcast(
            data = .,
            formula = domain ~ is.current,
            fun.aggregate = mean,
            value.var = "value"
          ) %>%
          ReplaceNames(., current.names = c("0","1"), new.names = c("Baseline","2018-2019")) %>%
          mutate(
            Trend = .[,3] - .[,2]
          ) %>%
          TransposeTable(., keep.first.colname = FALSE)
        
        tab3.state.avg.current.vs.baseline %<>%
          apply(
            X = tab3.state.avg.current.vs.baseline[,2:ncol(tab3.state.avg.current.vs.baseline)], 
            MARGIN = 2, 
            FUN = as.numeric
          ) %>%
          cbind(
            tab3.state.avg.current.vs.baseline[,1],
            .
          ) %>%
          ReplaceNames(., "Var.1", "")
        
        tab3.state.avg.current.vs.baseline[1,1] <- "Baseline"
  
  #DISTRICT/BUIDLING TABLES ----  
        
  ###                          ###    
  ### LOOP "c" BY REPORT UNIT  ###
  ###                          ###
  
    #Loop outputs
      tables.ls <- list()
    
    #Loop Measurement - progress bar & timing
      progress.bar.c <- txtProgressBar(min = 0, max = 100, style = 3)
      nrows.c <- lapply(config.tables.ls, nrow) %>% unlist
      c.loop.startime <- Sys.time()
    
    #Building Level Order
      building.level.order.v <- c("Elem.","High","Middle","Technology Ctr.","Other")
        
    #c <- 1 #LOOP TESTER 
    for(c in 1:length(unit.ids.sample)){   #START OF LOOP BY unit.id
      #Loop Prep----
        #Loop timing
          tic("REPORT LOOP DURATION",c)
        
        #Loop Inputs (both graphs and tables)
          unit.id.c <- unit.ids.sample[c]
        
        #Print loop messages
          if(c == 1){print("Forming tables for export...")}
          print("REPORT LOOP")
          print(paste("Loop #: ", c, " - Pct. Complete: ", 100*c/length(unit.ids.sample),sep = ""))
          print(paste("Unit id: ", unit.ids.sample[c]))
      
      #Tabs 1 & 2 ----
       
        #Loop Inputs
          config.tables.tab12.input.tb <- 
            config.tables.ls[[c]] %>% 
            filter(!is.na(table.type.id)) %>%
            filter(grepl("1|2", tab.type.id)) %>%
            filter(!is.state.table) %>%
            OrderDfByVar(., order.by.varname = "tab.type.id", rev = FALSE) %>%
            as_tibble()
          
          tables.tab12.ls <- list()
        
        #d <- 4
        for(d in 1:nrow(config.tables.tab12.input.tb)){ ### START OF LOOP "d" BY TABLE ###
          
          #Loop timing
            #tic("Tabs 1 & 2 loop iteration:", d)
          
          #Print loop messages
            print(paste("TABS 1 & 2 LOOP - Loop #: ", d, " - Pct. Complete: ", 100*d/nrow(config.tables.tab12.input.tb), sep = ""))
          
          #Define table configs for loop
            config.tables.tb.d <- config.tables.tab12.input.tb[d,]
          
          #CREATE TABLE
            if(config.tables.tb.d$tab.type.id == 1){
              
              #Define table aggregation formula
                table.formula.d <-
                  DefineTableRowColFormula(
                    row.header.varnames = strsplit(config.tables.tb.d$row.header.varname, ",") %>% unlist %>% as.vector,
                    col.header.varnames = strsplit(config.tables.tb.d$col.header.varname, ",") %>% unlist %>% as.vector
                  )

              #Define table source data
                filter.varnames.d <- config.tables.tb.d$filter.varname %>% strsplit(., ";") %>% unlist %>% as.vector
                filter.values.d <- config.tables.tb.d$filter.values %>% strsplit(., ";") %>% unlist %>% as.vector
                
                is.state.table <- ifelse(!"unit.id" %in% filter.varnames.d, TRUE, FALSE)
                is.domain.table <- grepl("domain", c(filter.varnames.d, table.formula.d)) %>% any
                
                #STATE data with NO domains in formula
                  if(is.state.table & !is.domain.table){table.source.data <- resp.full.nosplit.tb}
                
                #STATE data WITH domains in formula
                  if(is.state.table & is.domain.table){table.source.data <- resp.full.split.tb}
                
                #NON-STATE data NO domains in formula
                  if(!is.state.table & !is.domain.table){table.source.data <- resp.sample.nosplit.tb}
                
                #NON-STATE data WITH domains in formula
                  if(!is.state.table & is.domain.table){table.source.data <- resp.sample.split.tb}
                
              #Define table filtering vector
                table.filter.v <-
                  DefineTableFilterVector(
                    tb = table.source.data,
                    filter.varnames = filter.varnames.d,
                    filter.values = filter.values.d
                  )

              #Form final data frame
                if(table.filter.v %>% not %>% all){
                  table.d <- ""
                }else{
                  table.d <-  
                  table.source.data %>%
                  filter(table.filter.v) %>%
                  dcast(
                    ., 
                    formula = table.formula.d, 
                    value.var = config.tables.tb.d$value.varname, 
                    fun.aggregate = table.aggregation.function
                  ) %>%
                  .[,names(.)!= "NA"]
                

                  #Modifications for specific tables
                    if(grepl("building.level", table.formula.d) %>% any){
                      table.d <- 
                        left_join(
                          building.level.order.v %>% as.data.frame %>% ReplaceNames(., ".", "building.level"), 
                          table.d,
                          by = "building.level"
                        )
                    }
                    
                    if(!config.tables.tb.d$row.header){  #when don't want row labels
                      table.d <- table.d %>% select(names(table.d)[-1])
                    }
                }
            }
          
            if(config.tables.tb.d$tab.type.id == 2){
              table.d <- 
                tables.tab12.ls[
                  tables.tab12.ls %>%
                  lapply(
                    .,
                    function(x){
                      (
                        x$configs$tab.type.id %>%
                          unlist %>% as.vector %>%
                          equals(1)
                      ) &
                        (
                          x$configs$table.type.id %>%
                            unlist %>% as.vector() %>%
                            equals(config.tables.tb.d$table.type.id)
                        )
                    }
                  ) %>%
                  unlist %>% as.vector
                ] %>%
                .[[1]] %>% 
                .[["table"]]
            }
  
          #Table Storage
            table.d.storage.index <- length(tables.tab12.ls) %>% add(1)
            tables.tab12.ls[[table.d.storage.index]] <- list()
            tables.tab12.ls[[table.d.storage.index]]$configs <- config.tables.tb.d
            tables.tab12.ls[[table.d.storage.index]]$table <- table.d
          
          #tic.log(format = TRUE)
          #toc(log = TRUE, quiet = TRUE)
          
        } ### END OF LOOP "d" BY TABLE ###
        
        #Finalizing loop outputs  
          names(tables.tab12.ls) <- 
            paste(
              rep(unit.id.c, length(tables.tab12.ls)),
              ".",
              c(1:length(tables.tab12.ls)),
              ".",
              config.tables.tb.d$table.type.name, 
              sep = ""
            ) %>%
            gsub("-", " ", .) %>%
            gsub(" ", ".", .) %>%
            SubRepeatedCharWithSingleChar(., ".") %>%
            tolower
        
      #Tab 3 ----
        #Print status
          print("Tab 3 calculations begun...")
          
        #Loop Inputs
          config.tab3.tb <-   
            config.tables.ls[[c]] %>% 
            filter(!is.na(table.type.id)) %>%
            filter(grepl("3", tab.type.id)) %>%
            OrderDfByVar(., order.by.varname = "table.type.id", rev = FALSE)
          
          tables.tab3.ls <- list()
        
        #State Average Table - Last School Year vs. Current
          tables.tab3.ls[[1]] <- list()
          tables.tab3.ls[[1]]$configs <- config.tab3.tb[1,]
          tables.tab3.ls[[1]]$table <- tab3.state.avg.current.vs.previous.school.year  
        
        #Building Average Table - Last School year vs. Current  
          tables.tab3.ls[[2]] <- list()
          tables.tab3.ls[[2]]$configs <- config.tab3.tb[2,]
          tab3.bldg.current.vs.previous.school.year <- 
            resp.sample.split.tb %>% 
            filter(
              unit.id == unit.id.c & 
              is.most.recent.or.current == 1
            )
          
          if(nrow(tab3.bldg.current.vs.previous.school.year) < 1){
            tab3.bldg.current.vs.previous.school.year <- "" 
          }else{
            tab3.bldg.current.vs.previous.school.year %<>%
              SplitColReshape.ToLong(
                df = ., 
                id.varname = "resp.id", 
                split.varname = "domain", 
                split.char = ","
              ) %>%
              as_tibble() %>%
              dcast(
                data = .,
                formula = building.name + domain ~ year,
                fun.aggregate = mean,
                value.var = "value"
              )
            
            if(ncol(tab3.bldg.current.vs.previous.school.year) < 4){
              tab3.bldg.current.vs.previous.school.year %<>%
                mutate(
                  `Prev. School Year` = NA,
                  Trend = NA
                )
            }
            
            if(ncol(tab3.bldg.current.vs.previous.school.year) == 4){
              tab3.bldg.current.vs.previous.school.year %<>%
                mutate(
                  Trend = .[,4] - .[,3]
                )
            }
            
            tab3.bldg.current.vs.previous.school.year %<>%
              melt(
                ., 
                id.vars = c("building.name","domain")
              ) %>%
              dcast(
                ., 
                formula = building.name + variable ~ domain
              )
            
            tab3.bldg.current.vs.previous.school.year$variable %<>%
              as.character %>%
              gsub("2017-2018", "Prev. School Year", .)
          }
          
          tables.tab3.ls[[2]]$table <- tab3.bldg.current.vs.previous.school.year
  
        #State Average Table - Baseline vs. Current
          tables.tab3.ls[[3]] <- list()
          tables.tab3.ls[[3]]$configs <- config.tab3.tb[3,]
          tables.tab3.ls[[3]]$table <- tab3.state.avg.current.vs.baseline
            
          
        #Building Average Table - Baseline year vs. Current  
          tables.tab3.ls[[4]] <- list()
          tables.tab3.ls[[4]]$configs <- config.tab3.tb[4,]
          tab3.bldg.current.vs.baseline <- 
            resp.sample.split.tb %>% 
            filter(
              unit.id == unit.id.c & 
                is.baseline.or.current == 1
            ) %>%
            SplitColReshape.ToLong(
              df = ., 
              id.varname = "resp.id", 
              split.varname = "domain", 
              split.char = ","
            ) %>%
            as_tibble() %>%
            dcast(
              data = .,
              formula = building.name + domain ~ year,
              fun.aggregate = mean,
              value.var = "value"
            )
          
          if(ncol(tab3.bldg.current.vs.baseline) < 4){
            tab3.bldg.current.vs.baseline %<>%
              mutate(
                `Prev. School Year` = NA,
                Trend = NA
              )
          }
          
          if(ncol(tab3.bldg.current.vs.baseline) == 4){
            tab3.bldg.current.vs.baseline %<>%
              mutate(
                Trend = .[,4] - .[,3]
              )
          }
          
          tab3.bldg.current.vs.baseline %<>%
            melt(
              ., 
              id.vars = c("building.name","domain")
            ) %>%
            dcast(
              ., 
              formula = building.name + variable ~ domain
            )
          
          tab3.bldg.current.vs.baseline$variable %<>%
            as.character %>%
            gsub("0000", "Baseline", .)
          
          tables.tab3.ls[[4]]$table <- tab3.bldg.current.vs.baseline
        
      #Tab 4+ (Building Summaries) ----
          
        #Loop timing
          #tic("Tab 4 duration:")  
          
        #Loop Inputs
          config.tables.tab4.input.tb <- 
            config.tables.ls[[c]] %>% 
            filter(!is.na(table.type.id)) %>%
            filter(tab.type.id %in% c(4)) %>%
            filter(!is.na(loop.id))
          tables.tab4.ls <- list()
          
        #e <- 3
        for(e in 1:nrow(config.tables.tab4.input.tb)){ ### START OF LOOP "e" BY TABLE ###
            
            #Print loop messages
              print(paste("TAB 4 LOOP - Loop #: ", e, " - Pct. Complete: ", 100*e/nrow(config.tables.tab4.input.tb), sep = ""))
            
            config.tables.tb.e <- config.tables.tab4.input.tb[e,]
            
            #Define table aggregation formula
              table.formula.e <-
                DefineTableRowColFormula(
                  row.header.varnames = strsplit(config.tables.tb.e$row.header.varname, ",") %>% unlist %>% as.vector,
                  col.header.varnames = strsplit(config.tables.tb.e$col.header.varname, ",") %>% unlist %>% as.vector
                )
              
            #Define table source data
              filter.varnames.e <- config.tables.tb.e$filter.varname %>% strsplit(., ";") %>% unlist %>% as.vector
              filter.values.e <- config.tables.tb.e$filter.values %>% strsplit(., ";") %>% unlist %>% as.vector
              
              is.state.table <- ifelse(!"unit.id" %in% filter.varnames.e, TRUE, FALSE)
              is.domain.table <- grepl("domain", c(filter.varnames.e, table.formula.e)) %>% any
              
              #STATE data with NO domains in formula
                if(is.state.table & !is.domain.table){table.source.data <- resp.full.nosplit.tb}
              
              #STATE data WITH domains in formula
                if(is.state.table & is.domain.table){table.source.data <- resp.full.split.tb}
              
              #NON-STATE data NO domains in formula
                if(!is.state.table & !is.domain.table){table.source.data <- resp.sample.nosplit.tb}
              
              #NON-STATE data WITH domains in formula
                if(!is.state.table & is.domain.table){table.source.data <- resp.sample.split.tb}
              
            #Define table filtering vector
              loop.id.e <- config.tables.tb.e$loop.id
              table.filter.v <-
                DefineTableFilterVector(
                  tb = table.source.data,
                  filter.varnames = filter.varnames.e,
                  filter.values = filter.values.e
                )
            
            #Create table itself
              if(table.filter.v %>% not %>% all){
                table.e <- ""
              }
              
              if(table.filter.v %>% any){
                table.e <-  
                  table.source.data %>%
                  filter(table.filter.v) %>%
                  dcast(
                    ., 
                    formula = table.formula.e, 
                    value.var = config.tables.tb.e$value.varname, 
                    fun.aggregate = table.aggregation.function
                  ) %>%
                  .[,names(.)!= ""]
                
                table.e$trend <- 
                  ifelse(
                    IsError(table.e[,ncol(table.e)] - table.e[,ncol(table.e)-1]),
                    NA,
                    table.e[,ncol(table.e)] - table.e[,ncol(table.e)-1]
                  )
              
              #Modifications for specific tables
                if(grepl("building.level", table.formula.e) %>% any){
                  table.e <- 
                    left_join(
                      building.level.order.v %>% as.data.frame %>% ReplaceNames(., ".", "building.level"), 
                      table.e,
                      by = "building.level"
                    )
                }
                
                if(!config.tables.tb.e$row.header){  #when don't want row labels
                  table.e <- table.e %>% select(names(table.e)[-1])
                }
              }
              
            #Table storage
              #if(table.e == ""){print(e)}
              tables.tab4.ls[[e]] <- list()
              tables.tab4.ls[[e]]$configs <- config.tables.tb.e
              tables.tab4.ls[[e]]$table <- table.e
            
          } ### END OF LOOP "e" BY TABLE ###
      
      #Assemble final outputs for district ----  
        tables.ls[[c]] <- c(tables.tab12.state.ls, tables.tab12.ls, tables.tab3.ls, tables.tab4.ls)
      
      toc(log = TRUE, quiet = TRUE)  
    } ### END OF LOOP "c" BY REPORT UNIT     
        
  #Section Clocking
    toc(log = TRUE, quiet = FALSE)
    Sys.time() - sections.all.starttime


# 5-EXPORT -----------------------------------------------
  #EXPORT SETUP: CLOCKING, LOAD FUNCTIONS, ESTABLISH OUTPUTS DIRECTORY ----
    #Code Clocking
      section6.starttime <- Sys.time()
      
    #Load Configs Functions
      setwd(rproj.dir)
      #source("6-powerpoints export functions.r")
      
    #Create Outputs Directory
      if(sample.print){
        outputs.dir <- 
          paste(
            outputs.parent.folder,
            gsub(":",".",Sys.time()), 
            sep = ""
          )
      }
      
      dir.create(
        outputs.dir,
        recursive = TRUE
      )
    
  #EXPORT OF LONG DATA (if in global configs) ----
    if(export.long.data %>% as.logical){
      
      setwd(outputs.dir) 
      
      data.output.filename <- 
         paste(
          "longdata_",
          gsub(":",".",Sys.time()),
          ".csv",
          sep = ""
        )
      
      write.csv(
        resp.long.tb,
        file = data.output.filename
      )
    
    }
  
  #EXPORT TO EXCEL REPORTS - LOOP 'h' BY REPORT UNIT ----
      #Remove all objects except those needed to print (patch on memory errors)
        #rm(
        #  list = 
        #    setdiff(
        #      ls(), 
        #      c("sections.all.starttime","RemoveNA","unit.ids.sample", "tables.ls","config.text.ls","source.tables.dir","outputs.dir","sample.print")
        #    )
        #)
      
      #Define function to write workbooks inside of loop h
        WriteDistrictWorkbook <- 
          function(
            district.tables.list,
            district.text.list,
            district.file.name
          ){
            config.tables.h <- district.tables.list %>% lapply(., `[[`, 1) %>% do.call(rbind, .)
            wb <- loadWorkbook(file.name.h, create = FALSE)
            setStyleAction(wb, XLC$"STYLE_ACTION.NONE")
            
            building.names.for.district <- config.tables.h %>% select(loop.id) %>% unlist %>% unique %>% RemoveNA
            
            #i <- 1 #LOOP TESTER
            for(i in 1:length(district.tables.list)){  
              
              #Loop inputs
                if(
                  (district.tables.list[[i]]$table %>% dim %>% length %>% equals(1)) && (district.tables.list[[i]]$table %>% equals(""))
                ){
                  table.i <- ""
                }else{
                  table.i <- district.tables.list[[i]]$table
                }
                configs.i <- district.tables.list[[i]]$configs
                
                if(configs.i$tab.type.id == 4){ #customize tab name if need be for building summaries
                  building.num <- configs.i$loop.id %>% unique %>% equals(building.names.for.district) %>% which
                  
                  configs.i$tab.name <- 
                    getSheets(wb) %>% 
                    .[grepl("Building Summary", .)] %>% 
                    .[building.num]
                }
                
              #Print loop messages
                #print(paste("Loop i #: ", i, " - Table: ", configs.i$table.type.name, sep = ""))
              
              #Write Worksheets
                writeWorksheet(
                  object = wb, 
                  data = table.i,
                  sheet = configs.i$tab.name,
                  startRow = configs.i$startrow,
                  startCol = configs.i$startcol,
                  header = configs.i$header,
                  rownames = configs.i$row.header
                )
              
            } # END OF LOOP 'i' BY TABLE
            
            #Delete any extra building summary tabs
              #building.summary.tabs.with.data <- 
              #  getSheets(wb) %>% 
              #  .[grepl("Building Summary", .)] %>% 
              #  assign("building.summary.tabs", ., pos = 1) %>%
              #  .[1:length(building.names.for.district)]
              
              #extra.building.summary.tabnames <-
              #  building.summary.tabs[!building.summary.tabs %in% building.summary.tabs.with.data]
              
              #for(k in 1:length(extra.building.summary.tabnames)){
              #  removeSheet(
              #    object = wb, 
              #    sheet = extra.building.summary.tabnames[k]
              #  )
              #}
            
            ###                           ###
            ###   LOOP "m" BY TEXT ITEM   ###
            ###                           ###
            
            config.text.h <- district.text.list
            
            #m = 2 #LOOP TESTER
            for(m in 1:nrow(config.text.h)){
              #print(paste("Loop m #:", m, " - Pct. Complete: ", 100*m/nrow(config.text.h), sep = ""))
              
              #Write tables to building worksheet
                writeWorksheet(
                  object = wb,
                  data = config.text.h$text.value[m],
                  sheet = config.text.h$tab.name[m],
                  startRow = config.text.h$row.num[m],
                  startCol = config.text.h$col.num[m],
                  header = FALSE,
                  rownames = FALSE
                )
              
            } #END OF LOOP 'm' BY TEXT ITEM
            
            saveWorkbook(wb)
            print(paste("WORKBOOK SAVED. File: ", file.name.h, " - Pct. complete: ", 100*h/length(unit.ids.sample), sep = ""))
          }   
      
   
    ###                          ###    
  # ### LOOP "h" BY REPORT UNIT  ###
    ###                          ###
    
    #Progress Bar
      print(districts.that.already.have.reports)
      print(unit.ids.sample)
      progress.bar.h <- txtProgressBar(min = 0, max = 100, style = 3)
      maxrow.h <- tables.ls %>% lengths %>% sum
      
    #h <- 1 #LOOP TESTER
    for(h in 1:length(unit.ids.sample)){ 
      
      unit.id.h <- unit.ids.sample[h]  
                    
      #Set up target file
        template.file <- 
          paste(
            source.tables.dir,
            "dashboard_template.xlsx",
            sep = ""
          )
        
        if(sample.print){
          
          file.name.h <- 
             paste(
              "district dashboard_",
              unit.id.h,
              "_",
              gsub(":",".",Sys.time()),
              ".xlsx",
              sep = ""
            )
          
        }else{
           file.name.h <- 
            paste(
              "district dashboard_",
              unit.id.h,
              ".xlsx",
              sep = ""
            )
        }
        
        target.path.h <- 
            paste(
              outputs.dir,
              "\\",
              file.name.h,
              sep=""
            ) 
        
        file.copy(template.file, target.path.h)
        
        print(file.name.h)
      
  
      ###                       ###
  #   ###   LOOP "i" BY TABLE   ###
      ###                       ###
      
         
      setwd(outputs.dir)
      
      WriteDistrictWorkbook(
        district.tables.list = tables.ls[[h]],
        district.text.list = config.text.ls[[h]],
        district.file.name = file.name.h
      )

    } # END OF LOOP 'h' BY REPORT UNIT

# 6-WRAP UP -----------------------------------------------
  #Code Clocking
    code.runtime <- Sys.time() %>% subtract(sections.all.starttime) %>% round(., 2)
      
    print(
      paste(
        "Total Runtime for ", 
        length(unit.ids.sample), 
        " reports: ", 
        code.runtime,
        sep = ""
      )
    )
  
    #implied.total.runtime.for.all.reports <- 
    #  resp.full.nosplit.tb$unit.id %>% 
    #  unique %>% length %>% 
    #  divide_by(length(unit.ids.sample)) %>% 
    #  multiply_by(code.runtime)
    
    #print(
    #  paste(
    #    "Implied total runtime for all reports (min): ",
    #    implied.total.runtime.for.all.reports,
    #    sep = ""
    #  )
    #)
    
#SIGNAL CODE IS FINISHED BY OPENING A NEW WINDOW      
  windows()
    