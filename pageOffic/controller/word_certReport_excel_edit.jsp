<%@ page language="java" import="java.util.*,com.zhuozhengsoft.pageoffice.*,com.zhuozhengsoft.pageoffice.excelwriter.*,net.yuanh.util.*" pageEncoding="gb2312"%>
<--设置page的语言，导入依赖包-->
<%@ taglib uri="http://java.pageoffice.cn" prefix="po"%>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<%
String err="";
//精确度、不确定度等级
String accuracy = "";
String detectUuid="";
String jsessionId="";
PageOfficeCtrl poCtrl=(PageOfficeCtrl)request.getAttribute("poCtrl");
if(UtilValidate.isEmpty(request.getAttribute("reportFilePath"))){
	err="<script>alert('文件地址没有发现,请联系管理员');</script>";
}else{
	// poCtrl.OfficeVendor = PageOffice.OfficeVendorType.WPSOffice;

	poCtrl.setOfficeVendor(OfficeVendorType.AutoSelect);
	//设置服务器页面
	poCtrl.setServerPage(request.getContextPath()+"/poserver.zz");
	poCtrl.setTitlebar(false); //隐藏标题栏
	poCtrl.setMenubar(false); //隐藏菜单栏
	poCtrl.setOfficeToolbars(true);//隐藏Office工具条
	//添加自定义按钮
	poCtrl.addCustomToolButton("保存","Save",1);
	poCtrl.addCustomToolButton("另存为","showDialogSave",1);
	poCtrl.addCustomToolButton("打印","wordToPrint",6);
	//poCtrl.addCustomToolButton("打印预览","printPreview",7);
	poCtrl.addCustomToolButton("全屏切换", "SwitchFullScreen()", 4);
	poCtrl.addCustomToolButton("关闭","Close",15);
	//设置保存页面
	String taskOrderItemId = (String)request.getAttribute("taskOrderItemId");
	detectUuid = (String)request.getAttribute("detectUuid");
    String checkDate = (String)request.getAttribute("checkDate");
	jsessionId = (String)request.getAttribute("jsessionId");
	poCtrl.setSaveFilePage("saveFileCert?detectUuids="+detectUuid+"&checkDate="+checkDate+"&jsessionId="+jsessionId);
	
	Workbook workBook = new Workbook();
	//定义Sheet对象，"Sheet1"是打开的Excel表单的名称
	Sheet originalSheet = workBook.openSheet("1");
	Sheet certSheet = workBook.openSheet("2");
	//证书号、委托送检单位,物品器具名称
	 String resultNo="";String custName="";String certCorpName="";String sampleName="";String specification="";String factoryCode="";String manuFacturer="";
	 String certCorpAddr="" ; String resultConclus="";String currTestDateYear="";String currTestDateMonth="";String currTestDateDay="";String endTestDateYear="";
	 String endTestDateMonth="";String endTestDateDay="";
	 String standardName_code="";String labAddr="";String temData="";String humData="";String otherData="";String origLogNo="";String orderType="";
	 String testDateCycle="";
	//从request中获取所需的数据
	if(UtilValidate.isNotEmpty(request.getAttribute("resultNo"))){
	    resultNo =(String)request.getAttribute("resultNo") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("custName"))){
		custName =(String)request.getAttribute("custName") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("certCorpName"))){
		certCorpName =(String)request.getAttribute("certCorpName") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("sampleName"))){
		sampleName =(String)request.getAttribute("sampleName") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("specification"))){
		specification =(String)request.getAttribute("specification") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("factoryCode"))){
		factoryCode =(String)request.getAttribute("factoryCode") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("resultConclus"))){
		resultConclus =(String)request.getAttribute("resultConclus") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("certCorpAddr"))){
		certCorpAddr =(String)request.getAttribute("certCorpAddr") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("manuFacturer"))){
		manuFacturer =(String)request.getAttribute("manuFacturer") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("currTestDateYear"))){
		currTestDateYear =(String)request.getAttribute("currTestDateYear") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("currTestDateMonth"))){
		currTestDateMonth =(String)request.getAttribute("currTestDateMonth") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("currTestDateDay"))){
		currTestDateDay =(String)request.getAttribute("currTestDateDay") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("endTestDateYear"))){
		endTestDateYear =(String)request.getAttribute("endTestDateYear") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("endTestDateMonth"))){
		endTestDateMonth =(String)request.getAttribute("endTestDateMonth") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("endTestDateDay"))){
		endTestDateDay =(String)request.getAttribute("endTestDateDay") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("origLogNo"))){
		origLogNo =(String)request.getAttribute("origLogNo") ;
	}
	if(UtilValidate.isNotEmpty(request.getAttribute("orderType"))){
		orderType =(String)request.getAttribute("orderType") ;
	}
	//获取鉴定周期
	if(UtilValidate.isNotEmpty(request.getAttribute("testDateCycle"))){
		testDateCycle =(String)request.getAttribute("testDateCycle") ;
	}
	System.err.println("/fileInfo/"+request.getAttribute("reportFilePath")+"======================1111111111111111111");
	//第二页
	//检验依据（规程）
	if(UtilValidate.isNotEmpty(request.getAttribute("stdList"))){
        List<String> stdList = (List<String>) request.getAttribute("stdList");
			for(int i=0;i<stdList.size();i++){
				if(i==(stdList.size()-1)){
					standardName_code=standardName_code+stdList.get(i);
				}else{
					standardName_code=standardName_code+stdList.get(i)+"\n";
				}
			}
    }
	//测试地点
	if(UtilValidate.isNotEmpty(request.getAttribute("labAddr"))){
	    labAddr =(String)request.getAttribute("labAddr") ;
	}
	//温度
	if(UtilValidate.isNotEmpty(request.getAttribute("temData"))){
	    temData =(String)request.getAttribute("temData") ;
	}
	//湿度
	if(UtilValidate.isNotEmpty(request.getAttribute("humData"))){
	    humData =(String)request.getAttribute("humData") ;
	}
	//其他
	if(UtilValidate.isNotEmpty(request.getAttribute("otherData"))){
	    otherData =(String)request.getAttribute("otherData") ;
	}else{
		otherData="/";
	}

	 Cell cell = null;
    //测量审核
    if(UtilValidate.isNotEmpty(orderType)&&"order_type_3".equals(orderType)){
        String certTester = "";String certCorpCode = ""; String certPostCode = "";String certLinkMan = "";String certCorpTel = "";String certCorpFax = "";String labApprNo = "";
        if(UtilValidate.isNotEmpty(request.getAttribute("certTester"))){
			certTester = (String)request.getAttribute("certTester");
		}
		if(UtilValidate.isNotEmpty(request.getAttribute("certCorpCode"))){
			certCorpCode = (String)request.getAttribute("certCorpCode");
		}
		if(UtilValidate.isNotEmpty(request.getAttribute("certPostCode"))){
			certPostCode = (String)request.getAttribute("certPostCode");
		}
		if(UtilValidate.isNotEmpty(request.getAttribute("certLinkMan"))){
			certLinkMan = (String)request.getAttribute("certLinkMan");
		}
		if(UtilValidate.isNotEmpty(request.getAttribute("certCorpTel"))){
			certCorpTel = (String)request.getAttribute("certCorpTel");
		}
		if(UtilValidate.isNotEmpty(request.getAttribute("certCorpFax"))){
			certCorpFax = (String)request.getAttribute("certCorpFax");
		}
		if(UtilValidate.isNotEmpty(request.getAttribute("labApprNo"))){
			labApprNo = (String)request.getAttribute("labApprNo");
		}
		//设置原始记录基本信息位置
		String[] cellStr ={"E3","E4","R8","E8","E9","E10","Q10","E11","E5","E7","Q7","E6"};
		String[] valueStr ={certCorpName,certCorpAddr,currTestDateYear+"-"+currTestDateMonth+"-"+currTestDateDay,
				checkDate,resultConclus,temData,humData,certTester,sampleName,specification,
				factoryCode,manuFacturer};
		//原始记录插入基本信息
		for(int i=0;i<cellStr.length;i++){
			//获取Excel指定坐标位置
			cell = originalSheet.openCell(cellStr[i]);
			//插入数据
			cell.setValue(valueStr[i]);
		}
		//设置证书基本信息位置
		String[] certStr ={"X2","H3","AI3","H4","AF4","H5","AF5","C9","H9","AK12","H13"};
		System.err.println(certCorpAddr+"=============================");
		String[] certValueStr ={resultNo,certCorpName,labApprNo,certCorpAddr+" "+certPostCode,certCorpCode,certCorpTel+" "+certCorpFax,
				certLinkMan,sampleName,specification,currTestDateYear+"-"+currTestDateMonth+"-"+currTestDateDay,standardName_code};
		for(int i=0;i<certStr.length;i++){
			cell = certSheet.openCell(certStr[i]);
			cell.setValue(certValueStr[i]);
		}
	}else{
		 //设置原始记录基本信息位置
		String[] cellStr ={"E5","E6","P5"};
		String[] valueStr ={currTestDateYear+"-"+currTestDateMonth+"-"+currTestDateDay,
				endTestDateYear+"-"+endTestDateMonth+"-"+endTestDateDay,testDateCycle};
		//原始记录插入基本信息
		for(int i=0;i<cellStr.length;i++){
			cell = originalSheet.openCell(cellStr[i]);
			cell.setValue(valueStr[i]);
		}
		//设置证书基本信息位置
		String[] certStr ={"M16","I20","I24","I26","R26","I28","I22","D72","I76","Q76","Y76","K80"};
		String[] certValueStr ={resultNo,certCorpName,sampleName,specification,factoryCode,manuFacturer,certCorpAddr,
				labAddr,temData,humData,otherData,standardName_code};
		for(int i=0;i<certStr.length;i++){
			cell = certSheet.openCell(certStr[i]);
			cell.setValue(certValueStr[i]);
		}


        /*
        证书模板类型
        CONCLUS_TYPE_1	校准
        CONCLUS_TYPE_2	测试
        CONCLUS_TYPE_4	检测
        CONCLUS_TYPE_7	校准CNAS
        CONCLUS_TYPE_8	检测CNAS

        CONCLUS_TYPE_3	检定

        CONCLUS_TYPE_5	测量审核

        CONCLUS_TYPE_6	检定结果通知单
        */
		//检定类型证书插入发证日期，有效日期
        if(UtilValidate.isNotEmpty(request.getAttribute("resultTypeId")) && ("CONCLUS_TYPE_3".equals(request.getAttribute("resultTypeId")))) {
            String[] dateStr ={"Q37","U37","Y37","Q40","U40","Y40"};
            String[] dateValueStr ={currTestDateYear,currTestDateMonth,currTestDateDay,endTestDateYear,endTestDateMonth,endTestDateDay};
            for (int i = 0; i < dateStr.length; i++) {
                cell = certSheet.openCell(dateStr[i]);
                cell.setValue(dateValueStr[i]);
            }
        }else if(UtilValidate.isNotEmpty(request.getAttribute("resultTypeId")) && ("CONCLUS_TYPE_6".equals(request.getAttribute("resultTypeId")))){
        	//检定结果通知书类型
            String[] dateStr ={"Q37","U37","Y37"};
            String[] dateValueStr ={currTestDateYear,currTestDateMonth,currTestDateDay};
            for(int i=0;i<dateStr.length;i++){
                cell = certSheet.openCell(dateStr[i]);
                cell.setValue(dateValueStr[i]);
            }
        }else if(UtilValidate.isNotEmpty(request.getAttribute("resultTypeId")) && !"CONCLUS_TYPE_5".equals(request.getAttribute("resultTypeId"))){
            //证书类型为1,2,4,7,8的校准证书，插入校准日期
            String[] dateStr ={"H47","L47","P47"};
            String[] dateValueStr ={currTestDateYear,currTestDateMonth,currTestDateDay};
            for(int i=0;i<dateStr.length;i++){
                cell = certSheet.openCell(dateStr[i]);
                cell.setValue(dateValueStr[i]);
            }
        }
	
        //CANS证书添加接收日期
        if(UtilValidate.isNotEmpty(request.getAttribute("resultTypeId")) && ("CONCLUS_TYPE_7".equals(request.getAttribute("resultTypeId")) || "CONCLUS_TYPE_8".equals(request.getAttribute("resultTypeId")))){
        	String checkDateYear = "";
        	String checkDateMonth = "";
        	String checkDateDay = "";
        	if(UtilValidate.isNotEmpty(request.getAttribute("checkDateYear"))){
        		checkDateYear = (String)request.getAttribute("checkDateYear");
    		}
        	if(UtilValidate.isNotEmpty(request.getAttribute("checkDateMonth"))){
        		checkDateMonth = (String)request.getAttribute("checkDateMonth");
    		}
        	if(UtilValidate.isNotEmpty(request.getAttribute("checkDateDay"))){
        		checkDateDay = (String)request.getAttribute("checkDateDay");
    		}	
        	//证书类型为7,8的cans证书，插入接收日期
            String[] dateStr ={"H45","L45","P45"};
            String[] dateValueStr ={checkDateYear,checkDateMonth,checkDateDay};
            for(int i=0;i<dateStr.length;i++){
                cell = certSheet.openCell(dateStr[i]);
                cell.setValue(dateValueStr[i]);
            }
        }


        //校准的、检测、测试-----以前的代码---2019-02-16修改
		/*if(UtilValidate.isNotEmpty(request.getAttribute("resultTypeId")) &&
				("CONCLUS_TYPE_1".equals(request.getAttribute("resultTypeId"))||"CONCLUS_TYPE_4".equals(request.getAttribute("resultTypeId"))||"CONCLUS_TYPE_2".equals(request.getAttribute("resultTypeId")))){
			String[] dateStr ={"H47","L47","P47"};
			String[] dateValueStr ={currTestDateYear,currTestDateMonth,currTestDateDay};
			for(int i=0;i<dateStr.length;i++){
				cell = certSheet.openCell(dateStr[i]);
				cell.setValue(dateValueStr[i]);
			}
			//检定结果通知书
		}else if(UtilValidate.isNotEmpty(request.getAttribute("resultTypeId")) && ("CONCLUS_TYPE_6".equals(request.getAttribute("resultTypeId")))){
			String[] dateStr ={"Q37","U37","Y37"};
			String[] dateValueStr ={currTestDateYear,currTestDateMonth,currTestDateDay};
			for(int i=0;i<dateStr.length;i++){
				cell = certSheet.openCell(dateStr[i]);
				cell.setValue(dateValueStr[i]);
			}
		}else {
            String[] dateStr ={"Q37","U37","Y37","Q40","U40","Y40"};
            String[] dateValueStr ={currTestDateYear,currTestDateMonth,currTestDateDay,endTestDateYear,endTestDateMonth,endTestDateDay};
            for(int i=0;i<dateStr.length;i++){
                cell = certSheet.openCell(dateStr[i]);
                cell.setValue(dateValueStr[i]);
            }
        }*/
	}

	String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort();
	request.setAttribute("basePath", basePath);

	poCtrl.setWriter(workBook);
	//打开Word文档后事件
//	poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
	//打开Word文档
	poCtrl.setJsFunction_AfterDocumentOpened("insertImg");
	
	poCtrl.webOpen(basePath+"/fileInfo/"+request.getAttribute("reportFilePath"), OpenModeType.xlsNormalEdit, "张佚名");//打开文件路径，文件类型，当前打开文件用户


	poCtrl.setTagId("PageOfficeCtrl1");//此行必需
}
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>
<head>
<meta charset="utf-8">
<title>在线编辑报告</title>
   <script type="text/javascript" src="/js/pageofficeExt.js"></script>
   <script type="text/javascript" src="/js/jquery-1.9.1.min.js"></script>
   <script type="text/javascript" src="pageoffice.js" id="po_js_main"></script>
<script type="text/javascript">
	/*pageoffice页面最大化*/
    $(document).ready(function () {
        document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
	})

function SwitchFullScreen() {
    document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
}

function insertImg() {
    /**  
      *document.getElementById("PageOfficeCtrl1").InsertWebImage( ImageURL, Transparent, Position );
      *ImageURL  字符串类型，图片的路径。 
      *Transparent  布尔类型，可选参数，图片是否透明。默认值：FALSE，图片不透明；TRUE表示图片透明。注意：透明色为白色。
      *Position  整数类型，可选参数，浮于文字上方还是下方。默认值：4，图片浮于文字上方。 5，表示图片衬于文字下方。
      */
     //该方法默认插入图片到当前光标处，如果想插入到文档指定位置，可以在文档中插入一个书签来设置位置，然后先定位光标到书签处再插入图片
     //插入校准员签名
   //插入证书条码的信息
   //插入二维码信息
     var path =  "";//图片路径
     var basePath = "${basePath}";
     var orderType = "${orderType}";
     var cell = "";//书签
	//测量审核，插入检定员签名，不需证书条码和二维码
	if(orderType == "order_type_3"){
        cell = "AJ11";
        path =  "${certTester_excel}";
        locateBookMark(cell);
        document.getElementById("PageOfficeCtrl1").InsertWebImage(basePath+"/fileInfo/downLoadFileByPath?filePath="+path, false, 4);
	}else{
	   for (var i = 0; i < 2; i++) {
			if(i == 0){
				cell = "U2";
				path = "${imagePath_zsbhtm_excel}";
			}else if(i == 1){
				cell = "W52";
				path = "${imagePath_zsewm_excel}";
			}
			 if (path != null && path != '') {
				locateBookMark(cell);
				document.getElementById("PageOfficeCtrl1").InsertWebImage(basePath+"/fileInfo/downLoadFileByPath?filePath="+path, false, 4);
			 }
		}
	}
  }
  
  //定位书签到光标处
   function locateBookMark(cell) {
   //将光标定位到书签所在的位置
   var mac = "Function myfunc()" + " \r\n"
           + "    Range(\""+cell+"\").Select\n"
           // + "    Range(\"B4:D6\").Select\n"
           +"End Function";
   document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
}   
    function Save() {
	   	 document.getElementById("PageOfficeCtrl1").WebSave();
		 var reportResult =  document.getElementById("PageOfficeCtrl1").CustomSaveResult;
	   	if(reportResult == 'ok' ){
	        	var basePath = "${basePath}";
				var resultUuid = "${resultUuid}";
				$.ajax({
					url : basePath+"/resultInfo/updateResultInfo",
					type : 'POST',
					async:false,
					dataType:'json',
					contentType:"application/json;charset=UTF-8",
					data:JSON.stringify({resultUuid:resultUuid,isCertUpload:"Y"}),
					error : function(msg) {
					},
					success : function(data) {
					}
				});
				var resultNoView = $("#resultNo").val();
				window.external.CallParentFunc("pageOfficeValue('Y','"+resultNoView+"')")
	       		window.external.close();//执行完毕后关闭窗口
       }else{
       	alert("操作失败!");
   	}
   }
   function Close(){
	   window.external.close();//执行完毕后关闭窗口
   }
   function wordToPrint(){
	   document.getElementById("PageOfficeCtrl1").ShowDialog(4); 
	}
   function showDialogSave(){
	   document.getElementById("PageOfficeCtrl1").ShowDialog(3); 
   }
   function printPreview(){
	   document.getElementById("PageOfficeCtrl1").PrintPreview(); 
   }
   function AfterDocumentOpened() {
	   DeleteImage(30,22);//签发员
	   DeleteImage(33,22);//核验员
	  if($("#isTesterId2Show").val()!=null && $("#isTesterId2Show").val()!='undefined' && $("#isTesterId2Show").val()!='' && $("#isTesterId2Show").val()=='Y'){
		  
	  } else{
		  DeleteImage(36,22);//检定员
	  }
	   DeleteImage(34,1);//二维码
	   DeleteImage(4,17);//一维码
   }
   function updateCellAccuracy() {
		var htmlStyle="<style>*{margin:0;padding:0;}body{text-align: center;}";
		htmlStyle += "</style>";
		var headContent="<html><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"+htmlStyle+"</head><body>";
		var footerContent="</body></html>";
		var mainContent=document.getElementById("clientFileContent").innerHTML;// 获得html代码
		mainContent=mainContent.replace(/\"/g,"'").replace(/\n/g,"").replace(/\r/g,"");
		var strFileContent=headContent+mainContent+footerContent;
		var strBookmarkName="PO_equipTable";//书签名称
		PO_setValueTable(strFileContent);//插入表格数据
		PO_setValueTableCell(strBookmarkName,1,5,3);//在指定书签处插入表格中的某列某单元格的内容
	}
   $(function(){
       var num=$(document).height()-20;
       $("#main").css("height",num+"px");
   })
</script>
</head>
<body style="overflow-y:hidden;margin:0px;padding:0px;">
    <div style="width:auto; height:1024px;" id="main">
    	<input type="hidden" name="isToNext" value=""/>
			<%=poCtrl.getHtmlCode("PageOfficeCtrl1")%>
    </div>
    <input type="hidden" name="isToNext" id="isToNext" value="" />
    <input type="hidden" name="resultNo" id="resultNo" value="<%=request.getAttribute("resultNo") %>" />
    <input type="hidden"  id="isTesterId2Show" value="<%request.getAttribute("isTesterId2Show"); %>" />
    <div id="clientFileContent" style="display:none;"><%=accuracy %></div>
<%=err %>
</body>
</html>