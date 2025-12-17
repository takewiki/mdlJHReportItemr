#' 判断记录是否存在-------
#'
#' @param erpToken ERP口令
#' @param FBillNo  单据编号
#'
#' @return 返回值
#' @export
#'
#' @examples
#' coa_IsNew()
coa_IsNew <- function(erpToken = 'C0426D23-1927-4314-8736-A74B2EF7A039', FBillNo = "XSCKD-102-20250521-0004") {
  sql = paste0("select F_RDS_COA_TEMPLATENUMBER  from rds_erp_coa_vw_sal_outStock_task
where FBillNo = '",FBillNo,"'
order by FDATE")
  
  data = tsda::sql_select2(token = erpToken,sql = sql)
  ncount = nrow(data)
  if (ncount) {
    res = TRUE
  }else{
    res = FALSE
  }
  return(res)
  
}

#' 通过销售出现单获取COA模板号
#'
#' @param erpToken ERP口令
#' @param FBillNo 单据编号
#'
#' @return 返回值
#' @export
#'
#' @examples
#' coa_GetTemplateNumber()
coa_GetTemplateNumber <- function(erpToken = 'C0426D23-1927-4314-8736-A74B2EF7A039', FBillNo = "XSCKD-102-20250521-0004") {
  sql = paste0("select F_RDS_COA_TEMPLATENUMBER  from rds_erp_coa_vw_sal_outStock_task
where FBillNo = '",FBillNo,"'
order by FDATE")
  data = tsda::sql_select2(token = erpToken,sql = sql)
  
  ncount = nrow(data)
  if (ncount) {
    res = data$F_RDS_COA_TEMPLATENUMBER[1]
    if (res == ''|res == ' '){
      res = NULL
    }
  }else{
    res = NULL
  }
  return(res)
  
}



#' 获取COA表体走向
#'
#' @param erpToken ERP口令
#' @param FTemplateNumber 单据编号
#'
#' @return 返回值
#' @export
#'
#' @examples
#' coa_GetBodyDirection()
coa_GetBodyDirection <- function(erpToken = 'C0426D23-1927-4314-8736-A74B2EF7A039', FTemplateNumber = "M021") {
  sql = paste0("select FBodyDirection from rds_t_TemplateList where FTemplateNumber= '",FTemplateNumber,"'
")
  res = tsda::sql_select2(token = erpToken,sql = sql)
  
  
  return(res)
  
}




#' 获取单据清单
#'
#' @param erpToken ERP 口令
#'
#' @return 返回值
#' @export
#'
#' @examples
#' coa_TaskToSync()
coa_SyncAll <- function(erpToken = 'C0426D23-1927-4314-8736-A74B2EF7A039',outputDir = getwd(), delete_localFiles = 0) {
  sql_erp=paste0(" exec rds_proc_ReportItem_update")
  
  res=tsda::sql_update2(token = erp_token,sql_str = sql_erp )
  print('ERP字段同步成功')
  sql = paste0("select FBillNo  from rds_erp_coa_vw_sal_outStock_task")
  data = tsda::sql_select2(token = erpToken,sql = sql)
  
  ncount = nrow(data)
  res = 0
  if(ncount){
    
    for (i in 1:ncount) {
      FBillNo = data$FBillNo[i]
      step = coa_pdf(erpToken = erpToken,FBillNo = FBillNo,outputDir = outputDir,delete_localFiles = delete_localFiles)
      
      res = res + step
      
    }
  }
  return (res)
  
}
#' 获取客户名称
#'
#' @param erpToken ERP口令
#' @param FBillNo 单据编号
#'
#' @return 返回值
#' @export
#'
#' @examples
#' coa_GetCustomerName()
coa_GetCustomerName <- function(erpToken = 'C0426D23-1927-4314-8736-A74B2EF7A039', FBillNo = "XSCKD-102-20250521-0004") {
  sql = paste0("select FCustomerName  from rds_erp_coa_vw_sal_outStock_task
where FBillNo = '",FBillNo,"'
order by FDATE")
  data = tsda::sql_select2(token = erpToken,sql = sql)
  ncount = nrow(data)
  if (ncount) {
    res = data$FCustomerName[1]
    res = gsub(" ","_",res)
    res = gsub("\\.","_",res)
    res = gsub(",","_",res)
    res = gsub("\\(","_",res)
    res = gsub("\\)","_",res)
    if (res == ''){
      res = NULL
    }
  }else{
    res = NULL
  }
  return(res)
  
}

#' 获取日期
#'
#' @param erpToken ERP口令
#' @param FBillNo 单据编号
#'
#' @return 返回值
#' @export
#'
#' @examples
#' coa_GetFDate()
coa_GetFDate <- function(erpToken = 'C0426D23-1927-4314-8736-A74B2EF7A039', FBillNo = "XSCKD-102-20250521-0004") {
  sql = paste0("select cast(fdate as date) as FDate  from rds_erp_coa_vw_sal_outStock_task
where FBillNo = '",FBillNo,"' ")
  res = tsda::sql_select2(token = erpToken,sql = sql)
  
  return(res)
  
}



#' 返回元数据信息
#'
#' @param erpToken ERP口令
#' @param FTemplateNumber 模板号
#' @param FCOATableType 数据类型
#'
#' @return 返回值
#' @export
#'
#' @examples
#' coa_meta()
coa_meta <- function(erpToken = 'C0426D23-1927-4314-8736-A74B2EF7A039',
                     FTemplateNumber ='M001_100_IMP' ,
                     FCOATableType ='billHead') {
  sql = paste0("select FName_ERP_en,FTableName,FCells from [rds_t_ReportConfiguration]
where   FTemplateNumber ='",FTemplateNumber,"' and FCOATableType ='",FCOATableType,"'
AND  FTableName like '%sal_outStock%'

               ")
  data = tsda::sql_select2(token = erpToken,sql = sql)
  return(data)
  
}
#' 将excel列转
#'
#' @param col_str 列名
#'
#' @return 返回值
#'
#' @examples
#' excel_col_to_num()
excel_col_to_num <- function(col_str) {
  col_str <- toupper(col_str)
  chars <- strsplit(col_str, "")[[1]]
  sum <- 0
  for (char in chars) {
    sum <- sum * 26 + which(LETTERS == char)
  }
  return(sum)
}

# 将Excel坐标转换为行列数字
#' 将数据左边进行处理
#'
#' @param coord 提供坐标
#'
#' @return 返回值
#' @export
#'
#' @examples
#' excel_coord_to_numeric()
excel_coord_to_numeric <- function(coord) {
  col_str <- gsub("[^A-Za-z]", "", coord)
  row_num <- as.integer(gsub("[^0-9]", "", coord))
  col_num <- excel_col_to_num(col_str)
  return(c(col = col_num, row = row_num))
}

#' 将销售出库单数据写入EXCEL并上传到OSS对象存储
#'
#' @param erpToken ERP口令
#' @param outputDir 输出地址
#' @param delete_localFiles 是否删除本地文件
#' @param FBillNo 单据编号
#'
#' @return 返回值
#' @import openxlsx
#' @export
#'
#' @examples
#' coa_pdf()
coa_pdf <-function (erpToken = 'C0426D23-1927-4314-8736-A74B2EF7A039', FBillNo = "XSCKD-100-20250523-0001",
                    outputDir = getwd(), delete_localFiles = 0)
{
  
  flag_new = coa_IsNew(erpToken = erpToken,FBillNo=FBillNo)
  print(flag_new)
  print(1)
  if(flag_new){
    #全新的数据，做进一步处理
    #获取模板号
    template_coa = coa_GetTemplateNumber(erpToken = erpToken,FBillNo = FBillNo)
    
    if(is.null(template_coa)){
      print(paste0("销售出库单",FBillNo,"模板号为空，请及时维护"))
      res = 0
    }else{
      print(2)
      #进一步处理
      meta_head = coa_meta(erpToken = erpToken ,FTemplateNumber = template_coa,FCOATableType = 'billHead')
      ncount_meta_head = nrow(meta_head)
      fields_head = paste0(meta_head$FName_ERP_en,collapse = " , ")
      table_head = meta_head$FTableName[1]
      sql_head = paste0("select  ",fields_head,"   from  ",table_head," where FBillNo  = '",FBillNo,"' ")
      
      data_head =  tsda::sql_select2(token = erpToken,sql = sql_head)
      
      ncount_head = nrow(data_head)
      meta_entry = coa_meta(erpToken = erpToken ,FTemplateNumber = template_coa,FCOATableType = 'billEntry')
      ncount_meta_entry = nrow(meta_entry)
      fields_entry = paste0(meta_entry$FName_ERP_en,collapse = " , ")
      table_entry = meta_entry$FTableName[1]
      
      sql_entry = paste0("select  ",fields_entry,"   from  ",table_entry," where FBillNo  = '",FBillNo,"' order by F_RDS_COA_Lot ")
      data_entry =  tsda::sql_select2(token = erpToken,sql = sql_entry)
      ncount_entry = nrow(data_entry)
      
      
      
      
      if(ncount_head){
        print(3)
        #表头存在数据
        if(ncount_entry){
          #表体存在数据，进行相应的数据处理
          #获取完整的模板文件
          templateFile = paste0(outputDir, "/www/COA/",template_coa, ".xlsx")
          print(templateFile)
          excel_file <- openxlsx::loadWorkbook(templateFile)
          #写入表头数据
          for ( i in 1:ncount_meta_head) {
            #针对数据处理处理
            field_head = meta_head$FName_ERP_en[i]
            cell_head  = meta_head$FCells[i]
            print(cell_head)
            cellData_head = as.character(data_head[1,field_head])
            
            
            
            print(cellData_head)
            cellIndex_head =excel_coord_to_numeric(cell_head)
            indexCol = cellIndex_head['col']
            indexRow = cellIndex_head['row']
            
            
            header_style <- createStyle(
              fontName = "Calibri",
              fontSize = 10,
              halign = "center",       # 水平居中
              valign = "center",       # 垂直居中
              
              
            )
            
            
            openxlsx::writeData(wb = excel_file, sheet = "Sheet1", x = cellData_head,
                                startCol = indexCol, startRow = indexRow,
                                #colNames = FALSE,
                                #borders = "all" ,
                                headerStyle = header_style )
            
            
            
          }
          
          
          
          print(4)
          #处理表体数据
          for (j in 1:ncount_meta_entry) {
            fields_entry = meta_entry$FName_ERP_en[j]
            cell_entry = meta_entry$FCells[j]
            print(cell_entry)
            cellIndex_entry = excel_coord_to_numeric(cell_entry)
            for (k in 1:ncount_entry) {
              cellData_entry = data_entry[k ,fields_entry]
              
              # print('表体数据写入')
              # print(cellData_entry)
              # print(cellIndex_entry['col'])
              # print(cellIndex_entry['row'] + k -1)
              # print(cellIndex_entry['col']+k-1)
              # print(cellIndex_entry['row'])
              cellIndex_entry = excel_coord_to_numeric(cell_entry)
              print(template_coa)
              BodyDirection_coa = coa_GetBodyDirection(erpToken = erpToken,FTemplateNumber = template_coa)
              print(BodyDirection_coa)
              if(grepl("横", BodyDirection_coa)){
                openxlsx::writeData(wb = excel_file, sheet = "Sheet1", x = cellData_entry,
                                    startCol = cellIndex_entry['col']+k-1,
                                    startRow = cellIndex_entry['row'] ,
                                    colNames = FALSE)
                
              }else{
                openxlsx::writeData(wb = excel_file, sheet = "Sheet1", x = cellData_entry,
                                    startCol = cellIndex_entry['col'],
                                    startRow = cellIndex_entry['row'] + k -1,
                                    colNames = FALSE)
                
                
                
              }
              
              
              
              
            }
            
            
            
            
            
          }
          
          
          #处理文件名生成EXCEL
          print(5)
          
          # 生成文件名
          FCumstoerName = coa_GetCustomerName(erpToken = erpToken,FBillNo = FBillNo)
          
          #outputFile = paste0("COA_",FBillNo, "_", FCumstoerName,".xlsx")
          # 在生成文件名时替换空格
          clean_name <- function(text) {
            text <- gsub("\\s+", "_", text)  # 将所有空格替换为下划线
            gsub('[\\\\/:*?"<>|]', "_", text)  # 替换其他非法字符
          }
          
          FCumstoerName_six <- clean_name(substr(FCumstoerName, 1, 6))
          FBillNo_productName <- clean_name(sub(".*@", "", FBillNo))
          FDate=coa_GetFDate(erpToken = erpToken,FBillNo = FBillNo)
          FBillNo_no <- sub("@.*", "", FBillNo)
          
          outputFile = paste0("COA_",FCumstoerName_six,"_",FBillNo_productName,"_",FDate,"_",FBillNo_no,".xlsx")
          outputFile <- gsub("[()]", "", outputFile)
          
          pdf_base_name = paste0("COA_",FCumstoerName_six, "_",FBillNo_productName,"_",FDate,"_",FBillNo_no,".pdf")
          
          pdf_base_name <- gsub("[()]", "", pdf_base_name)
          xlsx_file_name = paste0(outputDir, "/", outputFile)
          print(xlsx_file_name)
          pdf_full_name = paste0(outputDir, "/", pdf_base_name)
          
          
          
          
          
          saveWorkbook(excel_file, xlsx_file_name, overwrite = TRUE)
          #生成PDF
          
          cmd = paste0("libreoffice --headless --convert-to pdf --outdir ",
                       outputDir, "  ", xlsx_file_name)
          Sys.setenv(LD_LIBRARY_PATH = paste("/usr/lib/libreoffice/program",
                                             Sys.getenv("LD_LIBRARY_PATH"), sep = ":"))
          system(cmd)
          #上传文件到OSS
          #start line
          dms_token = "AB8B9239-24C7-4599-99A9-69E09FEAEF01"
          oss_token = "A56C6025-A189-49A1-BE38-813BFAF52EF5"
          type = "jhcoa"
          #上传excel
          Url_excel = mdlOssr::rdOssFile_upload(dmsToken = dms_token,
                                                ossToken = oss_token, type = type, baseName = outputFile,
                                                fullName = xlsx_file_name)
          
          Url_pdf = mdlOssr::rdOssFile_upload(dmsToken = dms_token,
                                              ossToken = oss_token, type = type, baseName = pdf_base_name,
                                              fullName = pdf_full_name)
          
          
          
          # print("update")
          # print(FBillNo)
          
          
          sql_oss = paste0("update B set   F_RDS_QH_QualityReport =1,F_RDS_COA_XLSX ='",Url_excel,"',F_RDS_COA_PDF='",Url_pdf,"'
                            from t_sal_outStock a
				                    INNER JOIN T_sal_outStockENTRY B ON A.FID=B.FID
                            where CONCAT(a.FBillNo,'@',B.F_RDS_COA_ProductName,'_',F_RDS_COA_PageNumber) ='",FBillNo,"'")
          
          
          tsda::sql_update2(token = erpToken, sql_str = sql_oss)
          if (delete_localFiles) {
            if (file.exists(xlsx_file_name)) {
              file.remove(xlsx_file_name)
            }
            if (file.exists(pdf_full_name)) {
              file.remove(pdf_full_name)
            }
          }
          #end line
          
          res = 1
          
          
          
          
        }
        
      }
      else{
        res = 0
      }
      
      
      
      
      
    }
    
    
  }else{
    #任务不在COA任务中，不需要进行处理
    res = 0
  }
  
  return (res)
  
}

