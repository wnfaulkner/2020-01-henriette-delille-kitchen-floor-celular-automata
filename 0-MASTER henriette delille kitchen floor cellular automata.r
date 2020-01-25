#0000000000000000000000000000000000000000000000000000000000000000000000000#
#      	2020-01 henriette delille kitchen floor cellular automata   	    #
#0000000000000000000000000000000000000000000000000000000000000000000000000#

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
          wd <- paste(base.dir, "\\Dropbox\\5. REAL ESTATE\\2016-06 1219 Henriette Delille\\1. Maintenance, Repairs, Improvements\\2019-01 - General Renovation\\1. Design\\Kitchen Floor Cellular Automata\\", sep = "")
          rproj.dir <- paste(base.dir, "\\Documents\\GIT PROJECTS\\2020-01-henriette-delille-kitchen-floor-celular-automata", sep = "")
      }else{
        #Thinkpad T470
          wd <- "X:\\Google Drive File Stream\\Dropbox\\5. REAL ESTATE\\2016-06 1219 Henriette Delille\\1. Maintenance, Repairs, Improvements\\2019-01 - General Renovation\\1. Design\\Kitchen Floor Cellular Automata\\"
          rproj.dir <- paste(base.dir,"\\Documents\\GIT PROJECTS\\2020-01-henriette-delille-kitchen-floor-celular-automata", sep = "")
      }

    #Check Directories
      wd <- if(!dir.exists(wd)){choose.dir()}else{wd}
      rproj.dir <- if(!dir.exists(rproj.dir)){choose.dir()}else{rproj.dir}

    #Source Tables Directory (raw data, configs, etc.)
      #source.tables.dir <- paste(wd,"0. Inputs\\", sep = "")
      #if(dir.exists(source.tables.dir)){
      #  print("source.tables.dir exists.")
      #}else{
      #  print("source.tables.dir does not exist yet.")
      #}
      #print(source.tables.dir)

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
    source("99-Utils henriette delille kitchen floor cellular automata.r")

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
    library(gtools)

    #Section Clocking
      #section0.duration <- Sys.time() - section0.starttime
      #section0.duration

# 1- CREATE CELLULAR AUTOMATA -----------------------------------------

  #CONFIGS ----

    #Number of steps
      max.i <-
      8.7032 %>%
      divide_by(0.2032) %>% #longest distance in floor geometry (meters) divided by width of tiles (meters)
      ceiling %>%
      add(.*0.5) %>% #% cushion
      ceiling

    #Density Tolerance %
      tolerance = 0.1

    #Geometries to be subtracted from primary rectangle

      bath1 <- c(3.863/0.2302, 1.524/0.2302)
      #hvac <-

    #Rules for Testing (only interesting rules)
      rules.for.models <-
        paste(
          "r",
          c(110),
          #c(18, 22, 26, 30, 45, 73, 75, 86, 89, 105, 110, 118, 124, 135, 150, 169, 182),
          sep = ""
        )

  #INPUT TABLES ----

    #Complete Rules Table
      rules.base.tb <-
        permutations(
          n = 2,
          r = 3,
          v = c(0,1),
          repeats.allowed = T
        ) %>%
        .[order(rev(order(.[,1]))),] %>%
        as_tibble() %>%
        ReplaceNames(df = ., current.names = names(.), new.names = c("left","center","right"))


      rules.states.tb <-
        permutations(
          n = 2,
          r = 8,
          v = c(0,1),
          repeats.allowed = T
        ) %>%
        t %>%
        .[,-1] %>%
        as_tibble() %>%
        ReplaceNames(df = ., current.names = names(.), new.names = paste("r",1:ncol(.), sep = ""))

      rules.tb <-
        cbind(rules.states.tb, rules.base.tb)

      #testing to find column for rule 124 (confirm it is column 'r124')
        apply(rules.tb, 2, function(x){x %>% equals(c(0,1,1,1,1,1,0,0)) %>% all}) %>%
          as.vector %>%
          names(rules.tb)[.]

  #GENERATE AUTOMATA ----

    #Initial Matrix
      depth <- max.i
      width <- depth*2+5
      state.mtx <- matrix(0, nrow = depth, ncol = width)
      state.mtx[1,ceiling(width/2)] <- 1 #set initial seed
      state.tb <- state.mtx %>% as_tibble()

    i = 1
    #for(i in 1:length(rules.for.models)){ #START OF LOOP 'i' BY MODEL RULE

      #Rule for Model
        model.i <- rules.for.models[i]

      #State Change According to Rule Loop

        #state.ls <- list()
        progress.bar.j <- txtProgressBar(min = 0, max = 100, style = 3)
        max.j <- depth*width

        #j = 177 #loop tester
        for(j in 1:max.j){

          setTxtProgressBar(progress.bar.j, 100*j/max.j)

          row.j <- floor(j %>% divide_by(width) %>% subtract(10^-10)) %>% add(1)
          col.j <- j %>% subtract(row.j %>% subtract(1) %>% multiply_by(width))
          #if(col.j == 0)

          if(row.j == 1){ #no change if on first row
            state.tb[row.j, col.j] <- state.mtx[row.j, col.j]
            #state.ls[[j]] <- state.mtx[row.j, col.j]
            next()
          }

          if(col.j %in% c(1,2,width - 1, width)){ #mark state '0' if on two edge columns
            state.tb[row.j, col.j] <- 0
            #state.ls[[j]] <- 0
            next()
          }

          neighborhood.state.center <- state.tb[row.j - 1, col.j]
          neighborhood.state.left <- state.tb[row.j - 1, col.j - 1]
          neighborhood.state.right <- state.tb[row.j - 1, col.j + 1]
          neighborhood.v <- c(neighborhood.state.left, neighborhood.state.center, neighborhood.state.right) %>% unlist %>% as.vector

          state.filter.j <-
            rules.tb %>%
            select(left, center, right) %>%
            apply(., 1, function(x){x %>% equals(neighborhood.v) %>% all})

          state.tb[row.j, col.j] <-
            rules.tb %>%
            filter(state.filter.j) %>%
            select(model.i)

          #state.mtx[row.j, col.j] <- state.j

        }

        state.tb %>% as.data.frame() %>% .[1:10,floor((width/2)):100]

        #state.ls %>% length %>% equals(state.ls %>% unlist %>% length)
        #state.ls %>% lapply(., function(x){x %in% c(0,1)}) %>% unlist %>% not %>% which
        #state.ls %>% lapply(., function(x){is.null(x)}) %>% unlist %>% which

        #matrix(data = unlist(state.ls), nrow = depth, ncol = width)



  #TEST REGION DENSITY & GENERATE VISUAL ----

    dim.room <- c(7.3656, 4.6355) %>% divide_by(0.2032)
    max.startrow <- floor(nrow(state.tb) %>% subtract(dim.room[2]))

    designs.ls <- list()
    loop <- 100
    loop.permanent <- loop
    #i=1
    while(length(designs.ls) < 5){

      #Break if reached max number of attempts
        loop = loop-1
        if(loop == 0){
          print(
            paste(
              "Only ",
              length(designs.ls),
              " acceptable designs found in ",
              loop.permanent,
              " tries. Exiting loop"
            )
          )
          break
        }

      #Define primary rectangle
        startrow.i <- sample(1:max.startrow,1)

        startcol.i <-
          state.tb[startrow.i,] %>%
          unlist %>% as.vector %>%
          grep(1, .) %>%
          min %>%
          add(sample(1:(dim.room[1]/2), 1)*sample(c(1,-1),1))

        endrow.i <- ceiling(startrow.i + dim.room[2])
        endcol.i <- ceiling(startcol.i + dim.room[1])

        area.i <-
          state.tb[
            startrow.i:endrow.i,
            startcol.i:endcol.i
          ]

      #Subtract chunks for Bath 1, HVAC, counters

        #Bath 1
          area.i[
            floor(nrow(area.i) - bath1[2]):nrow(area.i),
            floor(ncol(area.i) - bath1[1]):ncol(area.i)
          ] <- NA

        #Counter
          #!

      density.i <- sum(area.i, na.rm = TRUE)/(nrow(area.i)*ncol(area.i))

      if(density.i > (0.5-tolerance) & density.i < (0.5+tolerance)){
        designs.ls.index <- length(designs.ls) + 1

        gt <- GridTopology(cellcentre=c(1,1),cellsize=c(1,1),cells=c(ncol(area.i), nrow(area.i))) %>% SpatialGrid()

        state.ls <- list()

        for(j in 1:nrow(area.i)){
          state.ls[[j]] <- area.i[j,] %>% unlist %>% as.vector
        }

        state.v <- do.call(c, state.ls) %>% as.data.frame

        sgdf = SpatialGridDataFrame(gt, state.v)

        designs.ls[[designs.ls.index]] <- "a"

        windows()
        image(sgdf, col=c("white", "black")) #!MAKE WITH 3 COLORS FOR SUBTRACTED GEOMETRIES

      }else{
        next()
      }
    }


    #} #END OF LOOP 'i' BY MODEL RULE







#---------------------------------------------------

#Original Code from: https://www.r-bloggers.com/cellular-automata-the-beauty-of-simplicity/



    #Cell state variable
    z <- data.frame(state=sample(0:0, width, replace=T))
    z[width/2, 1] <- 1
    z[width/2+1, 1] <- 1

  #Run loop to produce steps 1:depth
    for(i in (width+1):(width*depth)){

      ilf <- i-width-1
      iup <- i-width
      irg <- i-width+1

      if(i%%width==0){ irg <- i-2*width+1 }
      if(i%%width==1){ ilf <- i-1 }
      if(
        (z[ilf,1]+z[iup,1]+z[irg,1]>0)&(z[ilf,1]+z[iup,1]+z[irg,1]<3)
      ){
        st <- 1
      }else{
        st <- 0
      }

      nr <- as.data.frame(st)
      colnames(nr) <- c("state")
      z.output <- rbind(z,nr)

    }

    sgdf = SpatialGridDataFrame(gt, z)
    image(sgdf, col=c("white", "black"))

#----------------------------------------------------





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
