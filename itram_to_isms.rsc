//ISMS part 2
//ISMS Script Add-On to calculate External-Internal, Internal-External, and External-External traffic in an MPO model area

//This replaces the VB Macro process that was originally developed for the ISMS Prototype, and used in subsequent versions of ISMS.

//A. Pillman, 12/07/2021
//michael.pillman@iowadot.us
//515-239-1615

Macro "macro"

  RunMacro("TCB Init")
//Set-up Parameters
Scen_Directory = "C:\\TrafficModeling\\zResearch\\ISMS_2.0_Externals\\Testing\\Model\\2017\\"
mod_dir = "C:\\TrafficModeling\\zResearch\\ISMS_2.0_Externals\\Testing\\Model\\"
baseyr = "2017sasdfgasd
itram_trip = {"HBW", "HBO", "NHB", "LDW", "LDNW", "APRT", "SU", "COMBO"}
isms_trip = {"HBWL", "HBWM", "HBWH", "HBSC", "HBSH", "HBO", "NHB", "UNIV", "APRT", "HOSP", "RREC", "HOT", "SU", "COMBO"}
conv = tripgen + "itram_isms_conversion.bin"
periods = 3
//year_array = {base_parm, Horz_param} introudcing this when the new parameters for base and future MPO model years are intergrated into the GUI
years_array = {"2017", "2045"}
char_array = {"Base", "Future"}
scen_year = "2035"


//Copy PA Matrix file to Trip Generation as a working file
//create matrix currency you idjit
for z = 1 to 2 do //begin years array loop
  wrk_mtx = CopyFile(itram_subarea[z], tripgen + "PA_Subarea_"+char_array[z]+".mtx")

  //Copy Traffic Forecast file to Trip Generation as a working file
  Traf_table = Opeasdf
  strct[i] = strct[i] + {strct[i][1]}
  end

  strct = strct + {{"-- ","Character", 12, 3,,,,,,,, },
                      {"i_HBW","Real", 12, 3,,,,,,,, },{"i_HBO","Real", 12, 3,,,,,,,, },{"i_NHB","Real", 12, 3,,,,,,,, },
                      {"i_LDW","Real", 12, 3,,,,,,,, },{"i_LDNW","Real", 12, 3,,,,,,,, }, {"i_APRT","Real", 12, 3,,,,,,,, },
                      {"I_SU","Real", 12, 3,,,,,,,, }, {"i_COMBO","Real", 12, 3,,,,,,,, },
                    {"-- ","Character", 12, 3,,,,,,,, },
                      {"i_HBW_O"asdfsdafas
                      {"i_COMBO_O","Real", 16, 2,,,,,,,, },
                    {"-- ","Character", 12, 3,,,,,,,, },
                      {"i_HBW_D","Real", 12, 3,,,,,,,, },
                      {"i_HBO_D","Real", asdfasdfawsdfasd12, 3,,,,,,,, }, {"i_NHB_D","Real", 16, 2,,,,,,,, }, {"i_LDW_D","Real", 16, 2,,,,,,,, },
                      {"i_LDNW_D","Real", 16, 2,,,,,,,, }, {"i_APRT_D","Real", 16, 2,,,,,,,, }, {"i_SU_D","Real", 16, 2,,,,,,,, },
                      {"i_COMBO_D","Real", 16, 2,,,,,,,, },
                    {"-- ","Character", 12, 3,,,,,,,, },
                      {"PHBWL","Real", 16, 2,,,,,,,, }, {"PHBWM","Real", 16, 2,,,,,,,, },
                      {"PHBWH","Real", 16, 2,,,,,,,, }, {"PHBSC","Real", 16, 2,,,,,,,, }, {"PHBSH","Real", 16, 2,,,,,,,, },
                      {"PHBO","Real", 16, 2,,,,,,,, }, {"PNHB","Real", 16, 2,,,,,,,, }, {"PUNIV","Real", 16, 2,,,,,,,, },
                      {"PAPRT","Real", 16, 2,,,,,,,, }, {"PHOSP","Real", 16, 2,,,,,,,, }, {"PRREC","Real", 16, 2,,,,,,,, },
                      {"PHOT","Real", 16, 2,,,,,,,, }, {"PSU","Real", 16, 2,,,,,,,, }, {"PCOMBO","Real", 16, 2,,,,,,,, },
                    {"-- ","Character", 12, 3,,,,,,,, },
                      {"AHBWL","Real", 16, 2,,,,,,,, }, {"AHBWM","Real", 16, 2,,,,,,,, },
                      {"AHBWH","Real", 16, 2,,,,,,,, }, {"AHBSC","Real", 16, 2,,,,,,,, }, {"AHBSH","Real", 16, 2,,,,,,,, },
                      {"AHBO","Real", 16, 2,,,,,,,, }, {"ANHB","Real", 16, 2,,,,,,,, }, {"AUNIV","Real", 16, 2,,,,,,,, },
                      {"AAPRT","Real", 16, 2,,,,,,,, }, {"AHOSP","Real", 16, 2,,,,,,,, }, {"ARREC","Real", 16, 2,,,,,,,, },
                      {"AHOT","Real", 16, 2,,,,,,,, }, {"ASU","Real", 16, 2,,,,,,,, }, {"ACOMBO","Real", 16, 2,,,,,,,, },
                    {"-- ","Character", 12, 3,,,,,,,, },
                      {"Auto_SUM","Real", 16, 2,,,,,,,, }, {"Auto_HOLD","Real", 16, 2,,,,,,,, }, {"Auto_Factor","Real", 16, 2,,,,,,,, },
                      {"SU_SUM","Real", 16, 2,,,,,,,, }, {"SU_HOLD","Real", 16, 2,,,,,,,, }, {"SU_Factor","Real", 16, 2,,,,,,,, },
                      {"COMBO_SUM","Real", 16, 2,,,,,,,, }, {"COMBO_HOLD","Real", 16, 2,,,,,,,, }, {"COMBO_Factor","Real", 16, 2,,,,,,,, },
                    {"-- ","Character", 12, 3,,,,,,,, },
                      {"Group_Auto_Sum","Real", 16, 2,,,,,,,, }, {"Group_Auto_Factor","Real", 16, 2,,,,,,,, },
                      {"Group_SU_Sum","Real", 16, 2,,,,,,,, },{"Group_SU_Factor","Real", 16, 2,,,,,,,, },
                      {"Group_COMBO_Sum","Real", 16, 2,,,,,,,, },{"Group_COMBO_Factor","Real", 16, 2,,,,,,,, }}
  ModifyTable(conv_vw, strct)


  //Export subarea matrix trip purpose Origins and include Sum marginal
  for a = 1 to itram_trip.length do
  sub_origins = tripgen + "subarea_origins_"+itram_trip[a]+"_"+char_array[z]+".bin"
  m = OpenMatrix(tripgen + "pa_subarea_"+char_array[z]+".mtx", "True")
  mc = CreateMatrixCurrency(m, "Demand ("+itram_trip[a]+")",,,)
  ExportMatrix(mc, none ,"Rows", "FFB", sub_origins, {{"Marginal", "Sum"}}) // Get sum of Origins by taking the row sum of each itram production node


  //Join subarea matrix trip purpose origins to AB links
  subo = OpenTable("subarea_origins_"+itram_trip[a]+"_"+char_array[z], "FFB", {sub_origins})
  SetView(subo)
  SetView(conv_vw)
  join_origin = JoinViews("itram_to_isms_conversion+subarea_origins_"+itram_trip[a], conv_vw+".iTRAM_ID_AB", subo+".[Subarea Nodes]",{{"O"}})

  //Fill in iTRAM trip purpose origins
      get_subO_SUM = GetDataVector(join_origin+"|", subo+".Sum",)
      set_itram_O = SetDataVector(join_origin+"|", conv_vw+".i_"+itram_trip[a]+"_O", get_subO_SUM,) //Fill itram Origin fields

  end //matrix origins
  //Export subarea matrix trip purpose Destinations and include Sum marginal
  for a = 1 to itram_trip.length do
  sub_destinations = tripgen + "subarea_destinations_"+itram_trip[a]+"_"+char_array[z]+".bin"
  m = OpenMatrix(tripgen + "pa_subarea_"+char_array[z]+".mtx", "True")
  mc = CreateMatrixCurrency(m, "Demand ("+itram_trip[a]+")",,,)
  ExportMatrix(mc, none ,"Columns", "FFB", sub_destinations, {{"Marginal", "Sum"}}) // Get sum of destinations


  //Join subarea matrix trip purpose origins to AB links
  subd = OpenTable("subarea_destinations_"+itram_trip[a], "FFB", {sub_destinations})
  SetView(subd)
  SetView(conv_vw)
  join_origin = JoinViews("itram_to_isms_conversion+subarea_destinations_"+itram_trip[a], conv_vw+".iTRAM_ID_BA", subd+".[Subarea Nodes]",{{"O"}})

  //Fill in iTRAM trip purpose origins
      get_subd_SUM = GetDataVector(join_origin+"|", subd+".Sum",)
      set_itram_D = SetDataVector(join_origin+"|", conv_vw+".i_"+itram_trip[a]+"_D", get_subd_SUM,)

  end //matrix destinations
  //sum all O and Ds in a Sum All field
  for a = 1 to itram_trip.length do
    //Combine O and D
    //Fill Dataview
    Opts = null
    Opts.Input.[Dataview Set] = {conv, "itram_isms_conversion"+char_array[z]}
    Opts.Global.Fields = {"i_"+itram_trip[a]}
    Opts.Global.Method = "Formula"
    Opts.Global.Parameter = "i_"+itram_trip[a]+"_O +" + "i_"+itram_trip[a]+"_D"
    ret_value = RunMacro("TCB Run Operation", "Fill Dataview", Opts, &Ret)
    if !ret_value then goto quit
  end

      //SUM ALL TRIP PURPOSES FOR TRAFFIC FACTOR
      //Fill Dataview
      //Auto
      vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z])
      ok = (vw_set <> null)
      if !ok then goto quit

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "Auto_SUM", "nz(i_HBW) + nz(i_HBO) + nz(i_NHB) + nz(i_LDW) + nz(i_LDNW) + nz(i_APRT)", {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "Auto_SUM", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))

      //SU
      vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z])
      ok = (vw_set <> null)
      if !ok then goto quit

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "SU_SUM", "nz(i_SU)", {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "SU_SUM", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))

      //COMBO
      vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z])
      ok = (vw_set <> null)
      if !ok then goto quit

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "COMBO_SUM", "nz(i_COMBO)", {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "COMBO_SUM", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))

  //Create selection set of any TAZ that has a SUM of 0, fill SUM_All with traffic value and fill in HOLD with a value of 1
  CloseView(conv_vw)
  iTRAM_mode = {"Auto", "SU", "COMBO"}
  for a = 1 to iTRAM_mode.length do
    vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z])
        ok = (vw_set <> null)
        if !ok then goto quit

        {vw, set} = SplitString(vw_set)
        expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], iTRAM_mode[a]+"_HOLD", "if "+iTRAM_mode[a]+"_SUM = 0  then 1.00   else 0.00", {{"Type", "Integer"}, {"Width", 12}})
        vec = GetDataVector(vw_set, expr_fld,)
        SetDataVector(vw_set, iTRAM_mode[a]+"_HOLD", vec,)
        DestroyExpression(GetFieldFullSpec(vw, expr_fld))


     vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], iTRAM_mode[a]+"_HOLD", "Select * where "+iTRAM_mode[a]+"_HOLD = 1")
         ok = (vw_set <> null)
         if !ok then goto quit

         {vw, set} = SplitString(vw_set)
         expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], iTRAM_mode[a]+"_SUM", char_array[z]+"_"+iTRAM_mode[a]+"_Value", {{"Type", "Integer"}, {"Width", 12}})
         vec = GetDataVector(vw_set, expr_fld,)
         SetDataVector(vw_set, iTRAM_mode[a]+"_SUM", vec,)
         DestroyExpression(GetFieldFullSpec(vw, expr_fld))

         //0 Count
        check_zero = SelectByQuery(conv + "|", "Several", "Select * where "+char_array[z]+"_"+iTRAM_mode[a]+"_Value = 0")
        if check_zero = 0 then goto jump1
     vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], "ZeroReplace", "Select * where "+char_array[z]+"_"+iTRAM_mode[a]+"_Value = 0")
         ok = (vw_set <> null)
         if !ok then goto quit

         {vw, set} = SplitString(vw_set)
         expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], iTRAM_mode[a]+"_Factor", "1", {{"Type", "Real"}, {"Width", 12}})
         vec = GetDataVector(vw_set, expr_fld,)
         SetDataVector(vw_set, iTRAM_mode[a]+"_Factor", vec,)
         DestroyExpression(GetFieldFullSpec(vw, expr_fld))

        jump1:
     vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], "NormalFactor", "Select * where "+char_array[z]+"_"+iTRAM_mode[a]+"_Value <> 0")
         ok = (vw_set <> null)
         if !ok then goto quit

         {vw, set} = SplitString(vw_set)
         expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], iTRAM_mode[a]+"_Factor", char_array[z]+"_"+iTRAM_mode[a]+"_Value / "+iTRAM_mode[a]+"_SUM", {{"Type", "Real"}, {"Width", 12}})
         vec = GetDataVector(vw_set, expr_fld,)
         SetDataVector(vw_set, iTRAM_mode[a]+"_Factor", vec,)
         DestroyExpression(GetFieldFullSpec(vw, expr_fld))
  end

  //For iTRAM flow purposes = 0, do an equal distribution based off the counts

  for a = 1 to itram_trip.length do
    if a < 5 then do
      formula = "("+char_array[z]+"_Auto_Value / 2)/6"
      vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], "ZeroReplace", "Select * where Auto_HOLD = 1")
      ok = (vw_set <> null)
      if !ok then goto quit

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "i_"+itram_trip[a]+"_O", formula, {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "i_"+itram_trip[a]+"_O", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
  end //a <>5 or a<> 6
    if a = 5 then do
      formula = "("+char_array[z]+"_SU_Value / 2)"
      vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], "ZeroReplace", "Select * where SU_HOLD = 1")
      ok = (vw_set <> null)
      if !ok then goto quit

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "i_"+itram_trip[a]+"_O", formula, {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "i_"+itram_trip[a]+"_O", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
  end //a = 5
    if a = 6 then do
      formula = "("+char_array[z]+"_COMBO_Value / 2)"
      vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], "ZeroReplace", "Select * where COMBO_HOLD = 1")
      ok = (vw_set <> null)
      if !ok then goto quit

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "i_"+itram_trip[a]+"_O", formula, {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "i_"+itram_trip[a]+"_O", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
    end //a = 6
  end //itram_trip loop

  conv_vw = OpenTable("itram_isms_conversion"+char_array[z], "FFB", {conv})
  setview(conv_vw)

  for a = 1 to isms_trip.length do
//formulas for converting itram purpose values to daily ISMS purpose productions flow
    if a = 1 then formula = "(i_HBW_O + i_LDW_O)*HBWL*Auto_Factor"//HBWL
    if a = 2 then formula = "(i_HBW_O + i_LDW_O)*HBWM*Auto_Factor"//HBWM
    if a = 3 then formula = "(i_HBW_O + i_LDW_O)*HBWH*Auto_Factor"//HBWH
    if a = 4 then formula = "(i_HBO_O + i_LDNW_O)*HBSCP*Auto_Factor"//HBSC
    if a = 5 then formula = "(i_HBO_O + i_LDNW_O)*HBSHP*Auto_Factor"//HBSH
    if a = 6 then formula = "(i_HBO_O + i_LDNW_O)*HBOP*Auto_Factor"//HBO
    if a = 7 then formula = "i_NHB_O*Auto_Factor"//NHB
    if a = 8 then formula = "i_UNIV_O*Auto_Factor"//UNIV
    if a = 9 then formula = "i_APRT_O*Auto_Factor"//APRT
    if a = 10 then formula = "(i_HBO_O + i_LDNW_O)*HOSPP*Auto_Factor"//HOSP
    if a = 11 then formula = "(i_HBO_O + i_LDNW_O)*RRECP*Auto_Factor"//RREC
    if a = 12 then formula = "(i_HBO_O + i_LDNW_O)*HOTP*Auto_Factor"//HOT
    if a = 13 then formula = "I_SU_O*SU_Factor"//SU
    if a = 14 then formula = "i_COMBO_O*COMBO_Factor"//COMBO


    vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z])
        ok = (vw_set <> null)
        if !ok then goto quit

        {vw, set} = SplitString(vw_set)
        expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "P"+isms_trip[a], formula, {{"Type", "Real"}, {"Width", 12}})
        vec = GetDataVector(vw_set, expr_fld,)
        SetDataVector(vw_set, "P"+isms_trip[a], vec,)
        DestroyExpression(GetFieldFullSpec(vw, expr_fld))

  end

  for a = 1 to itram_trip.length do
    if a < 5 then do
      formula = "("+char_array[z]+"_Auto_Value / 2)/6"
      vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], "ZeroReplace", "Select * where Auto_HOLD = 1")
      ok = (vw_set <> null)
      if !ok then goto quit

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "i_"+itram_trip[a]+"_D", formula, {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "i_"+itram_trip[a]+"_D", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
  end //a <>5 or a<> 6
    if a = 5 then do
      formula = "("+char_array[z]+"_SU_Value / 2)"
      vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], "ZeroReplace", "Select * where SU_HOLD = 1")
      ok = (vw_set <> null)
      if !ok then goto quit

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "i_"+itram_trip[a]+"_D", formula, {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "i_"+itram_trip[a]+"_D", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
  end //a = 5
    if a = 6 then do
      formula = "("+char_array[z]+"_COMBO_COUNT / 2)"
      vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], "ZeroReplace", "Select * where COMBO_HOLD = 1")
      ok = (vw_set <> null)
      if !ok then goto quit

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "i_"+itram_trip[a]+"_D", formula, {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "i_"+itram_trip[a]+"_D", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
    end //a = 6
  end //itram_trip loop

  conv_vw = OpenTable("itram_isms_conversion"+char_array[z], "FFB", {conv})
  setview(conv_vw)

  for a = 1 to isms_trip.length do
//formulas for converting itram purpose values to daily ISMS purpose destinations flow
    if a = 1 then formula = "(i_HBW_D + i_LDW_D)*HBWL*Auto_Factor"//HBWL
    if a = 2 then formula = "(i_HBW_D + i_LDW_D)*HBWM*Auto_Factor"//HBWM
    if a = 3 then formula = "(i_HBW_D + i_LDW_D)*HBWH*Auto_Factor"//HBWH
    if a = 4 then formula = "(i_HBO_D + i_LDNW_D)*HBSCA*Auto_Factor"//HBSC
    if a = 5 then formula = "(i_HBO_D + i_LDNW_D)*HBSHA*Auto_Factor"//HBSH
    if a = 6 then formula = "(i_HBO_D + i_LDNW_D)*HBOA*Auto_Factor"//HBO
    if a = 7 then formula = "i_NHB_D*Auto_Factor"//NHB
    if a = 8 then formula = "i_UNIV_D*Auto_Factor"//UNIV
    if a = 9 then formula = "i_APRT_D*Auto_Factor"//APRT
    if a = 10 then formula = "(i_HBO_D + i_LDNW_D)*HOSPA*Auto_Factor"//HOSP
    if a = 11 then formula = "(i_HBO_D + i_LDNW_D)*RRECA*Auto_Factor"//RREC
    if a = 12 then formula = "(i_HBO_D + i_LDNW_D)*HOTA*Auto_Factor"//HOT
    if a = 13 then formula = "I_SU_D*SU_Factor"//SU
    if a = 14 then formula = "i_COMBO_D*COMBO_Factor"//COMBO


    vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z])
        ok = (vw_set <> null)
        if !ok then goto quit

        {vw, set} = SplitString(vw_set)
        expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "A"+isms_trip[a], formula, {{"Type", "Real"}, {"Width", 12}})
        vec = GetDataVector(vw_set, expr_fld,)
        SetDataVector(vw_set, "A"+isms_trip[a], vec,)
        DestroyExpression(GetFieldFullSpec(vw, expr_fld))
  end

  //Copy out itram_to_isms conversion to daily_pa_trips to mimic original paei files

  daily_pa = tripgen + "daily_pa_"+char_array[z]+".bin"
  wd_pa = tripgen + "wd_pa_"+char_array[z]+".bin"
  we_pa = tripgen + "we_pa_"+char_array[z]+".bin"
  wdwe = {"WD", "WE"}
  wdwe_file = {wd_pa, we_pa}


  conv_vw = OpenTable("itram_isms_conversion_"+char_array[z], "FFB", {conv})
  setview(conv_vw)
  exportview(conv_vw+"|", "FFB", daily_pa,{"TAZ",
    "PHBWL", "AHBWL", "PHBWM", "AHBWM", "PHBWH", "AHBWH", "PHBSC", "AHBSC", "PHBSH", "AHBSH", "PHBO", "AHBO", "PNHB", "ANHB",
    "PUNIV", "AUNIV", "PAPRT", "AAPRT", "PHOSP", "AHOSP", "PRREC", "ARREC", "PHOT", "AHOT", "PSU", "ASU", "PCOMBO", "ACOMBO"},)

    pa_vw = OpenTable("pa_trips_"+char_array[z], "FFB", {daily_pa})
    setview(pa_vw)

    strct = GetTableStructure(pa_vw)
    for i = 1 to strct.length do
    strct[i] = strct[i] + {strct[i][1]}
    end

    strct = strct + {{"-- ","Character", 12, 3,,,,,,,, },
                        {"HBWL","Real", 12, 3,,,,,,,, },{"HBWM","Real", 12, 3,,,,,,,, },{"HBWH","Real", 12, 3,,,,,,,, },
                        {"HBSC","Real", 12, 3,,,,,,,, },{"HBSH","Real", 12, 3,,,,,,,, },{"HBO","Real", 12, 3,,,,,,,, },
                        {"NHB","Real", 12, 3,,,,,,,, }, {"UNIV","Real", 12, 3,,,,,,,, }, {"APRT","Real", 12, 3,,,,,,,, },
                        {"HOSP","Real", 12, 3,,,,,,,, }, {"RREC","Real", 12, 3,,,,,,,, },{"HOT","Real", 12, 3,,,,,,,, },
                        {"SU","Real", 12, 3,,,,,,,, },{"COMBO","Real", 12, 3,,,,,,,, }}

  ModifyTable(pa_vw, strct)

//add together P and A purpose for total purpose value at station
  for a = 1 to isms_trip.length do
    vw_set = RunMacro("TCB Create View Set", daily_pa, "pa_trips_"+char_array[z])
        ok = (vw_set <> null)
        if !ok then goto quit

        {vw, set} = SplitString(vw_set)
        expr_fld = CreateExpression("pa_trips_"+char_array[z], "A"+isms_trip[a]+"+ P"+isms_trip[a], "A"+isms_trip[a]+" + P"+isms_trip[a], {{"Type", "Real"}, {"Width", 12}})
        vec = GetDataVector(vw_set, expr_fld,)
        SetDataVector(vw_set, isms_trip[a], vec,)
        DestroyExpression(GetFieldFullSpec(vw, expr_fld))

    end
CloseView(pa_vw)
  //Split daily file into WD/WE
  pa_vw = OpenTable("pa_trips_"+char_array[z], "FFB", {daily_pa})
  setview(pa_vw)
  Exportview(pa_vw+"|", "FFB", wd_pa,, )
  distro_dow_wd = ".75" //Value can be changed, thought being that 75% of traffic would be during the week compared to weekend.
  wd_formula = "* (("+distro_dow_wd+"*(5/7))/(("+distro_dow_wd+"*(5/7))+((1-"+distro_dow_wd+")*(2/7))))*(7/5)" //formula to convert to weekday portion of AADT
  we_formula = "*(((1-"+distro_dow_wd+")*(2/7))/(("+distro_dow_wd+"*(5/7))+((1-"+distro_dow_wd+")*(2/7))))*(7/2)" //formula to convert to weekend portion of AADT
  for a = 1 to isms_trip.length do
  vw_set = RunMacro("TCB Create View Set", wd_pa, "wd_pa_trips_"+char_array[z])
      ok = (vw_set <> null)
      if !ok then goto quit

      {vw, set} = SplitString(vw_set)
      expr_fldP = CreateExpression("wd_pa_trips_"+char_array[z], "P"+isms_trip[a], "P"+isms_trip[a]+wd_formula, {{"Type", "Real"}, {"Width", 12}})
      expr_fldA = CreateExpression("wd_pa_trips_"+char_array[z], "A"+isms_trip[a], "A"+isms_trip[a]+wd_formula, {{"Type", "Real"}, {"Width", 12}})
      vecP = GetDataVector(vw_set, expr_fldP,)
      vecA = GetDataVector(vw_set, expr_fldA,)
      SetDataVector(vw_set, "P"+isms_trip[a], vecP,)
      SetDataVector(vw_set, "A"+isms_trip[a], vecA,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fldP))
      DestroyExpression(GetFieldFullSpec(vw, expr_fldA))

  end //Conversion to Weekday

  pa_vw = OpenTable("pa_trips"+char_array[z], "FFB", {daily_pa})
  setview(pa_vw)
  Exportview(pa_vw+"|", "FFB", we_pa,, )
  //multiply fields by 2/7
  for a = 1 to isms_trip.length do
    vw_set = RunMacro("TCB Create View Set", we_pa, "we_pa_trips"+char_array[z])
        ok = (vw_set <> null)
        if !ok then goto quit

        {vw, set} = SplitString(vw_set)
        expr_fldP = CreateExpression("we_pa_trips"+char_array[z], "P"+isms_trip[a], "P"+isms_trip[a]+we_formula, {{"Type", "Real"}, {"Width", 12}})
        expr_fldA = CreateExpression("we_pa_trips"+char_array[z], "A"+isms_trip[a], "A"+isms_trip[a]+we_formula, {{"Type", "Real"}, {"Width", 12}})
        vecP = GetDataVector(vw_set, expr_fldP,)
        vecA = GetDataVector(vw_set, expr_fldA,)
        SetDataVector(vw_set, "P"+isms_trip[a], vecP,)
        SetDataVector(vw_set, "A"+isms_trip[a], vecA,)
        DestroyExpression(GetFieldFullSpec(vw, expr_fldP))
        DestroyExpression(GetFieldFullSpec(vw, expr_fldA))

  end // Conversion to Weekend

  //Split WD/WE files into time Periods
  if periods = 3 then do
    for b = 1 to 2 do //wdwe loop
      for a = 1 to 3 do //tod loop
        tp_split = tripgen + "paei_"+wdwe[b]+tod_char[a]+"_"+years_array[z]+".bin"
        wdwe_vw = OpenTable(wdwe[b]+"_pa_trips"+char_array[z], "FFB", {wdwe_file[b]})
        setview(wdwe_vw)
        Exportview(wdwe_vw+"|", "FFB", tp_split,, )
        for c = 1 to isms_trip.length do //isms trip purpose loop
          pformula3 = "P"+isms_trip[c]+"*"+tod_3_split[a]
          aformula3 = "A"+isms_trip[c]+"*"+tod_3_split[a]
          vw_set = RunMacro("TCB Create View Set", tp_split, "paei_"+wdwe[b]+tod_char[a]+"_"+years_array[z])
          ok = (vw_set <> null)
          if !ok then goto quit

          {vw, set} = SplitString(vw_set)
          expr_fldP = CreateExpression("paei_"+wdwe[b]+tod_char[a]+"_"+years_array[z], "P"+isms_trip[c], pformula3, {{"Type", "Real"}, {"Width", 12}})
          expr_fldA = CreateExpression("paei_"+wdwe[b]+tod_char[a]+"_"+years_array[z], "A"+isms_trip[c], aformula3, {{"Type", "Real"}, {"Width", 12}})
          vecP = GetDataVector(vw_set, expr_fldP,)
          vecA = GetDataVector(vw_set, expr_fldA,)
          SetDataVector(vw_set, "P"+isms_trip[a], vecP,)
          SetDataVector(vw_set, "A"+isms_trip[a], vecA,)
          DestroyExpression(GetFieldFullSpec(vw, expr_fldP))
          DestroyExpression(GetFieldFullSpec(vw, expr_fldA))
        end //isms trip purposes
      end //wdwe
    end //time period split
  end //if

  if periods = 4 then do
    for b = 1 to 2 do //wdwe loop
      for a = 1 to 4 do //tod loop
        tp_split = tripgen + "paei_"+wdwe[b]+tod_char[a]+"_"+years_array[z]+".bin"
        wdwe_vw = OpenTable(wdwe[b]+"_pa_trips", "FFB", {wdwe_file[b]})
        setview(wdwe_vw)
        Exportview(wdwe_vw+"|", "FFB", tp_split,, )
        for c = 1 to isms_trip.length do //isms trip purpose loop
          pformula4 = "P"+isms_trip[c]+"*"+tod_4_split[a]
          aformula4 = "A"+isms_trip[c]+"*"+tod_4_split[a]
          vw_set = RunMacro("TCB Create View Set", tp_split, "paei_"+wdwe[b]+tod_char[a]+"_"+year_array[z])
          ok = (vw_set <> null)
          if !ok then goto quit

          {vw, set} = SplitString(vw_set)
          expr_fldP = CreateExpression("paei_"+wdwe[b]+tod_char[a]+"_"+baseyr, "P"+isms_trip[c], pformula4, {{"Type", "Real"}, {"Width", 12}})
          expr_fldA = CreateExpression("paei_"+wdwe[b]+tod_char[a]+"_"+baseyr, "A"+isms_trip[c], aformula4, {{"Type", "Real"}, {"Width", 12}})
          vecP = GetDataVector(vw_set, expr_fldP,)
          vecA = GetDataVector(vw_set, expr_fldA,)
          SetDataVector(vw_set, "P"+isms_trip[a], vecP,)
          SetDataVector(vw_set, "A"+isms_trip[a], vecA,)
          DestroyExpression(GetFieldFullSpec(vw, expr_fldP))
          DestroyExpression(GetFieldFullSpec(vw, expr_fldA))
        end //isms trip purposes
      end //wdwe
    end //time period split
  end //if

  //Begin External-External process
  conv_vw = OpenTable("itram_isms_conversion"+char_array[z], "FFB", {conv})
  setview(conv_vw)

  //Get data on GROUP values for external station match
  vec = GetDataVector(conv_vw+"|", "GROUP",)
  x2 = SortArray({vec}, {{"Unique","True"}, {"Ascending","True"}})

  for a = 1 to x2[1].length do
    query = "Select * where GROUP = "+string(x2[1][a])
    vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], "Endorenna Utulien", query)
    ok = (vw_set <> null)
    if !ok then goto quit

    for b = 1 to iTRAM_mode.length do //sum group total value for factoring, fill value in factor field (repeats expected)
      subvec1 = GetDataVector(vw_set, char_array[z]+"_"+iTRAM_mode[b]+"_Value",)
      substat1 = VectorStatistic(subvec1, "Sum",)

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "Sinome Mullivan", string(substat1), {{"Type", "Integer"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "Group_"+iTRAM_mode[b]+"_Sum", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "pedo mellon a minno", char_array[z]+"_"+iTRAM_mode[b]+"_Value/Group_"+iTRAM_mode[b]+"_Sum", {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      SetDataVector(vw_set, "Group_"+iTRAM_mode[b]+"_Factor", vec,)
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))


      subvec2 = GetDataVector(vw_set, "Group_"+iTRAM_mode[b]+"_Sum",)
      substat2 = VectorStatistic(subvec2, "Sum",)
      if substat2 = 0 then do //if a mode has 0 as a total count, factor shall be 1
        vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z], "Melmelma noren sina", "Select * where Group_"+iTRAM_mode[b]+"_Sum = 0")
        ok = (vw_set <> null)
        if !ok then goto quit

        {vw, set} = SplitString(vw_set)
        expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], "Si dannatha nauva", "1", {{"Type", "Real"}, {"Width", 12}})
        vec = GetDataVector(vw_set, expr_fld,)
        SetDataVector(vw_set, "Group_"+iTRAM_mode[b]+"_Factor", vec,)
        DestroyExpression(GetFieldFullSpec(vw, expr_fld))
      end //Factor zero to 1 loop
    end //itram mode loop
  end //vector loop

//Create external-external matrix//
  external_trips = tripgen + "eetrips_"+years_array[z]+".mtx"
  m = OpenMatrix(tripgen + "pa_subarea_"+char_array[z]+".mtx", "True")
  mc = CreateMatrixCurrency(m, "Demand (HBW)", "Subarea Externals", "Subarea Externals",)
  nmc = CopyMatrix(mc, {{"File Name", external_trips}, {"Label", "eetrips_"+char_array[z]}, {"File Based", "Yes"}, {"Indices", "Current"}})

  vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z])
  ok = (vw_set <> null)
  if !ok then goto quit


  for a = 1 to isms_trip.length do
  //Add matrix cores
  m = OpenMatrix(tripgen + "eetrips_"+years_array[z]+".mtx", "True")
  matrix_cores = GetMatrixCoreNames(m)
  isms_trip_m = AddMatrixCore(m, isms_trip[a])

  mc1 = CreateMatrixCurrency(m, "Demand (HBW)",,,)
  mc2 = CreateMatrixCurrency(m, "Demand (HBO)",,,)
  mc3 = CreateMatrixCurrency(m, "Demand (NHB)",,,)
  mc4 = CreateMatrixCurrency(m, "Demand (LDW)",,,)
  mc5 = CreateMatrixCurrency(m, "Demand (LDNW)",,,)
  mc6 = CreateMatrixCurrency(m, "Demand (APRT)",,,)
  mc7 = CreateMatrixCurrency(m, "Demand (SU)",,,)
  mc8 = CreateMatrixCurrency(m, "Demand (COMBO)",,,)
  mc10 = CreateMatrixCurrency(m, isms_trip[a],,,)

//Get mean of splits from ISMS trip purposes to factor iTRAM values
    if a <= 3 then do
      matrix1 = mc1
      matrix2 = mc4
      vecHBW = GetDataVector(vw_set, isms_trip[a],)
      stat1 = VectorStatistic(vecHBW, "Mean",)

      mc10 := (matrix1+matrix2) * stat1
      goto jump
    end
    if isms_trip[a] = "HBSC" then do
      matrix1 = mc2
      matrix2 = mc5

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], isms_trip[a]+"average", "("+isms_trip[a]+"P+" + isms_trip[a]+"A) / 2", {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      isms_factor = VectorStatistic(vec, "Mean",)

      mc10 := (matrix1 + matrix2) * isms_factor
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
      goto jump
    end
    if isms_trip[a] = "HBSH" then do
      matrix1 = mc2
      matrix2 = mc5

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], isms_trip[a]+"average", "("+isms_trip[a]+"P+" + isms_trip[a]+"A) / 2", {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      isms_factor = VectorStatistic(vec, "Mean",)

      mc10 := (matrix1 + matrix2) * isms_factor
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
      goto jump
    end
    if isms_trip[a] = "HBO" then do
      matrix1 = mc2
      matrix2 = mc5

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], isms_trip[a]+"average", "("+isms_trip[a]+"P+" + isms_trip[a]+"A) / 2", {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      isms_factor = VectorStatistic(vec, "Mean",)

      mc10 := (matrix1 + matrix2) * isms_factor
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
      goto jump
    end
    if isms_trip[a] = "HOSP" then do
      matrix1 = mc2
      matrix2 = mc5

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], isms_trip[a]+"average", "("+isms_trip[a]+"P+" + isms_trip[a]+"A) / 2", {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      isms_factor = VectorStatistic(vec, "Mean",)

      mc10 := (matrix1 + matrix2) * isms_factor
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
      goto jump
    end
    if isms_trip[a] = "RREC" then do
      matrix1 = mc2
      matrix2 = mc5

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], isms_trip[a]+"average", "("+isms_trip[a]+"P+" + isms_trip[a]+"A) / 2", {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      isms_factor = VectorStatistic(vec, "Mean",)

      mc10 := (matrix1 + matrix2) * isms_factor
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
      goto jump
    end
    if isms_trip[a] = "HOT" then do
      matrix1 = mc2
      matrix2 = mc5

      {vw, set} = SplitString(vw_set)
      expr_fld = CreateExpression("itram_isms_conversion"+char_array[z], isms_trip[a]+"average", "("+isms_trip[a]+"P+" + isms_trip[a]+"A) / 2", {{"Type", "Real"}, {"Width", 12}})
      vec = GetDataVector(vw_set, expr_fld,)
      isms_factor = VectorStatistic(vec, "Mean",)

      mc10 := (matrix1 + matrix2) * isms_factor
      DestroyExpression(GetFieldFullSpec(vw, expr_fld))
      goto jump
    end
    if isms_trip[a] = "NHB" then do
      mc10 := mc3
      goto jump
    end

    if isms_trip[a] = "APRT" then do
      mc10 := mc6
      goto jump
    end

    if isms_trip[a] = "SU" then do
      mc10 := mc7
      goto jump
    end

    if isms_trip[a] = "COMBO" then do
      mc10 := mc8
      goto jump
    end

  jump:

  end // add matrix cores, compute conversions from iTRAM trip purposes to ISMS trip purposes

  //Create ISMS TAZ Production Index
  Opts = null
      Opts.Input.[Current Matrix] = external_trips
      Opts.Input.[Index Type] = "Row"
      Opts.Input.[View Set] = {conv, "itram_isms_conversion"+char_array[z]}
      Opts.Input.[Old ID Field] = {conv, "iTRAM_ID_AB"}
      Opts.Input.[New ID Field] = {conv, "TAZ"}
      Opts.Output.[New Index] = "TAZ-Production"

      ok = RunMacro("TCB Run Operation", "Add Matrix Index", Opts, &Ret)

  //Create ISMS TAZ Attraction Index
  Opts = null
      Opts.Input.[Current Matrix] = external_trips
      Opts.Input.[Index Type] = "Column"
      Opts.Input.[View Set] = {conv, "itram_isms_conversion"+char_array[z]}
      Opts.Input.[Old ID Field] = {conv, "iTRAM_ID_BA"}
      Opts.Input.[New ID Field] = {conv, "TAZ"}
      Opts.Output.[New Index] = "TAZ-Attraction"

      ok = RunMacro("TCB Run Operation", "Add Matrix Index", Opts, &Ret)

  conv_vw = OpenTable("itram_isms_conversion"+char_array[z], "FFB", {conv})
  setview(conv_vw)

  m = OpenMatrix(external_trips, "True")
  //Vector  multiply matrices for PAs
  for a = 1 to isms_trip.length do
    if a <= 12 then do
      mc = RunMacro("TCB Create Matrix Currency", external_trips, isms_trip[a], "TAZ-Production", "Subarea Externals") //set to multiply origins
      ok = (mc <> null)
      if !ok then goto quit


      mc1 = RunMacro("TCB Create Matrix Currency", external_trips, isms_trip[a], "Subarea Externals", "TAZ-Attraction") //set to multiply destinations
      ok = (mc1 <> null)
      if !ok then goto quit

      v_field = conv+".Auto_Factor"
    end

    if a = 13 then do
      mc = RunMacro("TCB Create Matrix Currency", external_trips, "SU", "TAZ-Production", "Subarea Externals") //set to multiply origins
      ok = (mc <> null)
      if !ok then goto quit
      v_field = conv+".SU_Factor"

      mc1 = RunMacro("TCB Create Matrix Currency", external_trips, "SU", "Subarea Externals", "TAZ-Attraction") //set to multiply destinations
      ok = (mc1 <> null)
      if !ok then goto quit

      v_field = conv+".SU_Factor"
    end

    if a = 14 then do
      mc = RunMacro("TCB Create Matrix Currency", external_trips, "COMBO", "TAZ-Production", "Subarea Externals") //set to multiply origins
      ok = (mc <> null)
      if !ok then goto quit

      mc1 = RunMacro("TCB Create Matrix Currency", external_trips, "COMBO", "Subarea Externals", "TAZ-Attraction") //set to multiply destinations
      ok = (mc1 <> null)
      if !ok then goto quit

      v_field = conv+".COMBO_Factor"
    end

  vw_set = RunMacro("TCB Create View Set", conv, "itram_isms_conversion"+char_array[z])
  ok = (vw_set <> null)
  if !ok then goto quit

    if a < 12 then do
      vec = RunMacro("TCB GetDataVector", vw_set, "Group_Auto_Factor", {{"Match Matrix Index", {mc, "TAZ", "Yes"}}})  // Match-matrix-index info array: {matrix_currency, id_field, apply_by_row}
      ok = (vec <> null)
      if !ok then goto quit

      vec1 = RunMacro("TCB GetDataVector", vw_set, "Group_Auto_Factor", {{"Match Matrix Index", {mc1, "TAZ", "No"}}})  // Match-matrix-index info array: {matrix_currency, id_field, apply_by_column}
      ok1 = (vec1 <> null)
      if !ok then goto quit
    end

    if a = 13 then do
      vec = RunMacro("TCB GetDataVector", vw_set, "Group_SU_Factor", {{"Match Matrix Index", {mc, "TAZ", "Yes"}}})  // Match-matrix-index info array: {matrix_currency, id_field, apply_by_row}
      ok = (vec <> null)
      if !ok then goto quit

      vec1 = RunMacro("TCB GetDataVector", vw_set, "Group_SU_Factor", {{"Match Matrix Index", {mc1, "TAZ", "No"}}})  // Match-matrix-index info array: {matrix_currency, id_field, apply_by_column}
      ok1 = (vec1 <> null)
      if !ok then goto quit
    end

    if a = 14 then do
      vec = RunMacro("TCB GetDataVector", vw_set, "Group_COMBO_Factor", {{"Match Matrix Index", {mc, "TAZ", "Yes"}}})  // Match-matrix-index info array: {matrix_currency, id_field, apply_by_row}
      ok = (vec <> null)
      if !ok then goto quit

      vec1 = RunMacro("TCB GetDataVector", vw_set, "Group_COMBO_Factor", {{"Match Matrix Index", {mc1, "TAZ", "No"}}})  // Match-matrix-index info array: {matrix_currency, id_field, apply_by_colum}
      ok1 = (vec1 <> null)
      if !ok then goto quit
    end

    mc := nz(mc * vec)
    mc1 := nz(mc1 * vec1)

  end //isms matrixVector

  //Split daily EE matrix into WD/WE
  m = OpenMatrix(tripgen + "eetrips_"+years_array[z]+".mtx", "True")
  mc = CreateMatrixCurrency(m, "Demand (HBW)",,,)
  weekdaymc = CopyMatrix(mc, {{"File Name", tripgen + "eeutrips_wd_"+years_array[z]+".mtx"}, {"Label", "eetrips_wd"+years_array[z]}, {"File Based", "Yes"}, {"Indices", "Current"}})
  weekendmc = CopyMatrix(mc, {{"File Name", tripgen + "eeutrips_we_"+years_array[z]+".mtx"}, {"Label", "eetrips_we"+years_array[z]}, {"File Based", "Yes"}, {"Indices", "Current"}})
  distro_dow_wd = .75
  wd_formula = ((distro_dow_wd*(5/7))/((distro_dow_wd*(5/7))+((1-distro_dow_wd)*(2/7))))*(7/5)
  we_formula = (((1-distro_dow_wd)*(2/7))/((distro_dow_wd*(5/7))+((1-distro_dow_wd)*(2/7))))*(7/2)

  for a = 1 to isms_trip.length do
    md = OpenMatrix(tripgen + "eeutrips_wd_"+years_array[z]+".mtx", "True")
    me = OpenMatrix(tripgen + "eeutrips_we_"+years_array[z]+".mtx", "True")
    wdmc = CreateMatrixCurrency(md, isms_trip[a],,,)
    wemc = CreateMatrixCurrency(me, isms_trip[a],,,)

     wdmc := wdmc * wd_formula //apply weekday formula to all isms trip purposes matrices
     wemc := wemc * we_formula //apply weekend formula to all isms trip purposes matrices

  end
  eemat = {wdmc, wemc}
  wdwe = {"WD", "WE"}
  //Split wd/we into AM/PM/OP/(MD if 4 time periods)
  if periods = 3 then do
    for a = 1 to 3 do
      for b = 1 to 2 do
        mcsplit = CopyMatrix(eemat[b], {{"File Name", tripgen + "eeutrips_"+wdwe[b]+tod_char[a]+"_"+years_array[z]+".mtx"},{"Label", "eeutrips_"+wdwe[b]+tod_char[a]+"_"+years_array[z]}, {"File Based", "Yes"}, {"Indices", "Current"}})
        tpmcpath = tripgen + "eeutrips_"+wdwe[b]+tod_char[a]+"_"+years_array[z]+".mtx"
        tpmc = OpenMatrix(tpmcpath, "True")
        for c = 1 to 14 do
          mc = CreateMatrixCurrency(tpmc, isms_trip[c],,, )
          mc := mc * StringToReal(tod_3_split[a])
        end
      end
    end
  end

  if periods = 4 then do
    for a = 1 to 4 do
      for b = 1 to 2 do
        mcsplit = CopyMatrix(eemat[b], {{"File Name", tripgen + "eeutrips_"+wdwe[b]+tod_char[a]+"_"+years_array[z]+".mtx"},{"Label", "eeutrips_"+wdwe[b]+tod_char[a]+"_"+years_array[z]}, {"File Based", "Yes"}, {"Indices", "Current"}})
        tpmc = tripgen + "eeutrips_"+wdwe[b]+tod_char[a]+"_"+baseyr+".mtx"
        tpmc = OpenMatrix(tpmcpath, "True")
        for c = 1 to 14 do
          mc = CreateMatrixCurrency(tpmc, isms_trip[c],,,)
          mc := mc * StringToReal(tod_4_split[a])


          end
        end
      end
  end
  CloseView(conv_vw)
  CloseView(pa_vw)
  RunMacro("G30 File Close All")
end //years array z = 1 to 2 do

///////temporary//////
quit:
    Return( RunMacro("TCB Closing", ok, True ) )


Return(ok)

EndMacro
