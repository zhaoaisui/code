<%@ page language="java" import="java.util.*,com.zhuozhengsoft.pageoffice.*,com.zhuozhengsoft.pageoffice.excelwriter.*,net.yuanh.util.*" pageEncoding="gb2312"%>
<--����page�����ԣ�����������-->
<%@ taglib uri="http://java.pageoffice.cn" prefix="po"%>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<%
String err="";
//��ȷ�ȡ���ȷ���ȵȼ�
String accuracy = "";
String detectUuid="";
String jsessionId="";
PageOfficeCtrl poCtrl=(PageOfficeCtrl)request.getAttribute("poCtrl");
if(UtilValidate.isEmpty(request.getAttribute("reportFilePath"))){
	err="<script>alert('�ļ���ַû�з���,����ϵ����Ա');</script>";
}else{
	// poCtrl.OfficeVendor = PageOffice.OfficeVendorType.WPSOffice;

	poCtrl.setOfficeVendor(OfficeVendorType.AutoSelect);
	//���÷�����ҳ��
	poCtrl.setServerPage(request.getContextPath()+"/poserver.zz");
	poCtrl.setTitlebar(false); //���ر�����
	poCtrl.setMenubar(false); //���ز˵���
	poCtrl.setOfficeToolbars(true);//����Office������
	//����Զ��尴ť
	poCtrl.addCustomToolButton("����","Save",1);
	poCtrl.addCustomToolButton("���Ϊ","showDialogSave",1);
	poCtrl.addCustomToolButton("��ӡ","wordToPrint",6);
	//poCtrl.addCustomToolButton("��ӡԤ��","printPreview",7);
	poCtrl.addCustomToolButton("ȫ���л�", "SwitchFullScreen()", 4);
	poCtrl.addCustomToolButton("�ر�","Close",15);
	//���ñ���ҳ��
	String taskOrderItemId = (String)request.getAttribute("taskOrderItemId");
	detectUuid = (String)request.getAttribute("detectUuid");
    String checkDate = (String)request.getAttribute("checkDate");
	jsessionId = (String)request.getAttribute("jsessionId");
	poCtrl.setSaveFilePage("saveFileCert?detectUuids="+detectUuid+"&checkDate="+checkDate+"&jsessionId="+jsessionId);
	
	Workbook workBook = new Workbook();
	//����Sheet����"Sheet1"�Ǵ򿪵�Excel��������
	Sheet originalSheet = workBook.openSheet("1");
	Sheet certSheet = workBook.openSheet("2");
	//֤��š�ί���ͼ쵥λ,��Ʒ��������
	 String resultNo="";String custName="";String certCorpName="";String sampleName="";String specification="";String factoryCode="";String manuFacturer="";
	 String certCorpAddr="" ; String resultConclus="";String currTestDateYear="";String currTestDateMonth="";String currTestDateDay="";String endTestDateYear="";
	 String endTestDateMonth="";String endTestDateDay="";
	 String standardName_code="";String labAddr="";String temData="";String humData="";String otherData="";String origLogNo="";String orderType="";
	 String testDateCycle="";
	//��request�л�ȡ���������
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
	//��ȡ��������
	if(UtilValidate.isNotEmpty(request.getAttribute("testDateCycle"))){
		testDateCycle =(String)request.getAttribute("testDateCycle") ;
	}
	System.err.println("/fileInfo/"+request.getAttribute("reportFilePath")+"======================1111111111111111111");
	//�ڶ�ҳ
	//�������ݣ���̣�
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
	//���Եص�
	if(UtilValidate.isNotEmpty(request.getAttribute("labAddr"))){
	    labAddr =(String)request.getAttribute("labAddr") ;
	}
	//�¶�
	if(UtilValidate.isNotEmpty(request.getAttribute("temData"))){
	    temData =(String)request.getAttribute("temData") ;
	}
	//ʪ��
	if(UtilValidate.isNotEmpty(request.getAttribute("humData"))){
	    humData =(String)request.getAttribute("humData") ;
	}
	//����
	if(UtilValidate.isNotEmpty(request.getAttribute("otherData"))){
	    otherData =(String)request.getAttribute("otherData") ;
	}else{
		otherData="/";
	}

	 Cell cell = null;
    //�������
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
		//����ԭʼ��¼������Ϣλ��
		String[] cellStr ={"E3","E4","R8","E8","E9","E10","Q10","E11","E5","E7","Q7","E6"};
		String[] valueStr ={certCorpName,certCorpAddr,currTestDateYear+"-"+currTestDateMonth+"-"+currTestDateDay,
				checkDate,resultConclus,temData,humData,certTester,sampleName,specification,
				factoryCode,manuFacturer};
		//ԭʼ��¼���������Ϣ
		for(int i=0;i<cellStr.length;i++){
			//��ȡExcelָ������λ��
			cell = originalSheet.openCell(cellStr[i]);
			//��������
			cell.setValue(valueStr[i]);
		}
		//����֤�������Ϣλ��
		String[] certStr ={"X2","H3","AI3","H4","AF4","H5","AF5","C9","H9","AK12","H13"};
		System.err.println(certCorpAddr+"=============================");
		String[] certValueStr ={resultNo,certCorpName,labApprNo,certCorpAddr+" "+certPostCode,certCorpCode,certCorpTel+" "+certCorpFax,
				certLinkMan,sampleName,specification,currTestDateYear+"-"+currTestDateMonth+"-"+currTestDateDay,standardName_code};
		for(int i=0;i<certStr.length;i++){
			cell = certSheet.openCell(certStr[i]);
			cell.setValue(certValueStr[i]);
		}
	}else{
		 //����ԭʼ��¼������Ϣλ��
		String[] cellStr ={"E5","E6","P5"};
		String[] valueStr ={currTestDateYear+"-"+currTestDateMonth+"-"+currTestDateDay,
				endTestDateYear+"-"+endTestDateMonth+"-"+endTestDateDay,testDateCycle};
		//ԭʼ��¼���������Ϣ
		for(int i=0;i<cellStr.length;i++){
			cell = originalSheet.openCell(cellStr[i]);
			cell.setValue(valueStr[i]);
		}
		//����֤�������Ϣλ��
		String[] certStr ={"M16","I20","I24","I26","R26","I28","I22","D72","I76","Q76","Y76","K80"};
		String[] certValueStr ={resultNo,certCorpName,sampleName,specification,factoryCode,manuFacturer,certCorpAddr,
				labAddr,temData,humData,otherData,standardName_code};
		for(int i=0;i<certStr.length;i++){
			cell = certSheet.openCell(certStr[i]);
			cell.setValue(certValueStr[i]);
		}


        /*
        ֤��ģ������
        CONCLUS_TYPE_1	У׼
        CONCLUS_TYPE_2	����
        CONCLUS_TYPE_4	���
        CONCLUS_TYPE_7	У׼CNAS
        CONCLUS_TYPE_8	���CNAS

        CONCLUS_TYPE_3	�춨

        CONCLUS_TYPE_5	�������

        CONCLUS_TYPE_6	�춨���֪ͨ��
        */
		//�춨����֤����뷢֤���ڣ���Ч����
        if(UtilValidate.isNotEmpty(request.getAttribute("resultTypeId")) && ("CONCLUS_TYPE_3".equals(request.getAttribute("resultTypeId")))) {
            String[] dateStr ={"Q37","U37","Y37","Q40","U40","Y40"};
            String[] dateValueStr ={currTestDateYear,currTestDateMonth,currTestDateDay,endTestDateYear,endTestDateMonth,endTestDateDay};
            for (int i = 0; i < dateStr.length; i++) {
                cell = certSheet.openCell(dateStr[i]);
                cell.setValue(dateValueStr[i]);
            }
        }else if(UtilValidate.isNotEmpty(request.getAttribute("resultTypeId")) && ("CONCLUS_TYPE_6".equals(request.getAttribute("resultTypeId")))){
        	//�춨���֪ͨ������
            String[] dateStr ={"Q37","U37","Y37"};
            String[] dateValueStr ={currTestDateYear,currTestDateMonth,currTestDateDay};
            for(int i=0;i<dateStr.length;i++){
                cell = certSheet.openCell(dateStr[i]);
                cell.setValue(dateValueStr[i]);
            }
        }else if(UtilValidate.isNotEmpty(request.getAttribute("resultTypeId")) && !"CONCLUS_TYPE_5".equals(request.getAttribute("resultTypeId"))){
            //֤������Ϊ1,2,4,7,8��У׼֤�飬����У׼����
            String[] dateStr ={"H47","L47","P47"};
            String[] dateValueStr ={currTestDateYear,currTestDateMonth,currTestDateDay};
            for(int i=0;i<dateStr.length;i++){
                cell = certSheet.openCell(dateStr[i]);
                cell.setValue(dateValueStr[i]);
            }
        }
	
        //CANS֤����ӽ�������
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
        	//֤������Ϊ7,8��cans֤�飬�����������
            String[] dateStr ={"H45","L45","P45"};
            String[] dateValueStr ={checkDateYear,checkDateMonth,checkDateDay};
            for(int i=0;i<dateStr.length;i++){
                cell = certSheet.openCell(dateStr[i]);
                cell.setValue(dateValueStr[i]);
            }
        }


        //У׼�ġ���⡢����-----��ǰ�Ĵ���---2019-02-16�޸�
		/*if(UtilValidate.isNotEmpty(request.getAttribute("resultTypeId")) &&
				("CONCLUS_TYPE_1".equals(request.getAttribute("resultTypeId"))||"CONCLUS_TYPE_4".equals(request.getAttribute("resultTypeId"))||"CONCLUS_TYPE_2".equals(request.getAttribute("resultTypeId")))){
			String[] dateStr ={"H47","L47","P47"};
			String[] dateValueStr ={currTestDateYear,currTestDateMonth,currTestDateDay};
			for(int i=0;i<dateStr.length;i++){
				cell = certSheet.openCell(dateStr[i]);
				cell.setValue(dateValueStr[i]);
			}
			//�춨���֪ͨ��
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
	//��Word�ĵ����¼�
//	poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
	//��Word�ĵ�
	poCtrl.setJsFunction_AfterDocumentOpened("insertImg");
	
	poCtrl.webOpen(basePath+"/fileInfo/"+request.getAttribute("reportFilePath"), OpenModeType.xlsNormalEdit, "������");//���ļ�·�����ļ����ͣ���ǰ���ļ��û�


	poCtrl.setTagId("PageOfficeCtrl1");//���б���
}
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>
<head>
<meta charset="utf-8">
<title>���߱༭����</title>
   <script type="text/javascript" src="/js/pageofficeExt.js"></script>
   <script type="text/javascript" src="/js/jquery-1.9.1.min.js"></script>
   <script type="text/javascript" src="pageoffice.js" id="po_js_main"></script>
<script type="text/javascript">
	/*pageofficeҳ�����*/
    $(document).ready(function () {
        document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
	})

function SwitchFullScreen() {
    document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
}

function insertImg() {
    /**  
      *document.getElementById("PageOfficeCtrl1").InsertWebImage( ImageURL, Transparent, Position );
      *ImageURL  �ַ������ͣ�ͼƬ��·���� 
      *Transparent  �������ͣ���ѡ������ͼƬ�Ƿ�͸����Ĭ��ֵ��FALSE��ͼƬ��͸����TRUE��ʾͼƬ͸����ע�⣺͸��ɫΪ��ɫ��
      *Position  �������ͣ���ѡ���������������Ϸ������·���Ĭ��ֵ��4��ͼƬ���������Ϸ��� 5����ʾͼƬ���������·���
      */
     //�÷���Ĭ�ϲ���ͼƬ����ǰ��괦���������뵽�ĵ�ָ��λ�ã��������ĵ��в���һ����ǩ������λ�ã�Ȼ���ȶ�λ��굽��ǩ���ٲ���ͼƬ
     //����У׼Աǩ��
   //����֤���������Ϣ
   //�����ά����Ϣ
     var path =  "";//ͼƬ·��
     var basePath = "${basePath}";
     var orderType = "${orderType}";
     var cell = "";//��ǩ
	//������ˣ�����춨Աǩ��������֤������Ͷ�ά��
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
  
  //��λ��ǩ����괦
   function locateBookMark(cell) {
   //����궨λ����ǩ���ڵ�λ��
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
	       		window.external.close();//ִ����Ϻ�رմ���
       }else{
       	alert("����ʧ��!");
   	}
   }
   function Close(){
	   window.external.close();//ִ����Ϻ�رմ���
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
	   DeleteImage(30,22);//ǩ��Ա
	   DeleteImage(33,22);//����Ա
	  if($("#isTesterId2Show").val()!=null && $("#isTesterId2Show").val()!='undefined' && $("#isTesterId2Show").val()!='' && $("#isTesterId2Show").val()=='Y'){
		  
	  } else{
		  DeleteImage(36,22);//�춨Ա
	  }
	   DeleteImage(34,1);//��ά��
	   DeleteImage(4,17);//һά��
   }
   function updateCellAccuracy() {
		var htmlStyle="<style>*{margin:0;padding:0;}body{text-align: center;}";
		htmlStyle += "</style>";
		var headContent="<html><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"+htmlStyle+"</head><body>";
		var footerContent="</body></html>";
		var mainContent=document.getElementById("clientFileContent").innerHTML;// ���html����
		mainContent=mainContent.replace(/\"/g,"'").replace(/\n/g,"").replace(/\r/g,"");
		var strFileContent=headContent+mainContent+footerContent;
		var strBookmarkName="PO_equipTable";//��ǩ����
		PO_setValueTable(strFileContent);//����������
		PO_setValueTableCell(strBookmarkName,1,5,3);//��ָ����ǩ���������е�ĳ��ĳ��Ԫ�������
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