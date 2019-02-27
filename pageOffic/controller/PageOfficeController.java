package net.yuanh.client.oms.controller;


import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.yuanh.thirdservice.oms.service.*;
import net.yuanh.util.*;
import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PDFCtrl;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;

import javolution.util.FastMap;
import net.yuanh.client.base.BaseController;
import net.yuanh.client.base.OmsStatusConstant;
import net.yuanh.thirdservice.ams.service.IEquipInfoService;
import net.yuanh.thirdservice.crm.service.ICustomerService;
import net.yuanh.thirdservice.dms.service.IFileConfigService;
import net.yuanh.thirdservice.dms.service.IFileInfoService;
import net.yuanh.thirdservice.dms.service.ITplOrgService;
import net.yuanh.thirdservice.snp.service.IStandardService;
import net.yuanh.thirdservice.ump.service.ISubcontracterService;
import net.yuanh.thirdservice.ump.service.IUserPersonService;

/**
 * @Title: PageOfficeController.java
 * @Package net.yuanh.client.oms.controller
 * @Description: 在线打开和保存原始记录和证书
 * @author fandongsheng
 * @date 2018年9月28日 上午9:39:52
 * @version V1.0
 */
@Controller
@RequestMapping(value = "/editPageOffice")
public class PageOfficeController extends BaseController {
	private final Logger logger = Logger.getLogger(PageOfficeController.class);

	@Autowired
	private IDetectInfoService detectInfoService;
	@Autowired
	private ICustomerService customerService;
	@Autowired
	private IStandardService standardService;
	@Autowired
	private IEquipInfoService equipInfoService;
	@Autowired
	private ISubcontracterService subcontracterService;
	@Autowired
	private IResultInfoService resultInfoService;
	@Autowired
	private IUserPersonService userPersonService;
	@Autowired
	private IFileInfoService fileInfoService;
	@Autowired
	private IFileConfigService fileConfigService;
	@Autowired
	private IWaitTestService waitTestService;
	@Autowired
	private OrderInfoController orderInfoController;
	@Autowired
	private ITplOrgService tplOrgService;
	@Autowired
	private IItemRequireService itemRequireService;

	// demo例子
	/*
	 * @RequestMapping(value="/onlineEditCert") public ModelAndView
	 * onlineEditCert(HttpServletRequest request){ PageOfficeCtrl poCtrl=new
	 * PageOfficeCtrl(request);
	 * poCtrl.setServerPage(request.getContextPath()+"/poserver.zz"); String
	 * detectUuids = request.getParameter("detectUuids"); String userUuid =
	 * request.getParameter("userUuid");
	 *
	 * //证书参数 String path = "d:\\test.doc"; request.setAttribute("poCtrl",
	 * poCtrl); request.setAttribute("path", path);
	 * poCtrl.webOpen(path,OpenModeType.docAdmin,""); ModelAndView mv = new
	 * ModelAndView("word"); return mv; }
	 */

	/**
	 * 在线编辑证书
	 *
	 * @param request
	 * @return
	 */
	@RequestMapping(value = "/onlineEditCert")
	public ModelAndView onlineEditCert(HttpServletRequest request) {
		//1，创建返回值对象，spring中的ModelAndView，封装要跳转的视图名称
		ModelAndView mv = new ModelAndView();
		//2，使用request创建pageoffice对象，并将所需要的数据封装到request中
		PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
		poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//此行必须有
		
		String detectUuids = request.getParameter("detectUuids");
		JSONObject jsonParam=new JSONObject();
		String jsessionId=request.getParameter("jsessionId");
		jsonParam.put("jsessionId", jsessionId);
		Map<String, Object> userInfo = getBaseDataToJSONObject(jsonParam);//获取到用户信息
		request.setAttribute("jsessionId", jsessionId);//返回给页面jsession
		String userUuid = String.valueOf(userInfo.get("userUuid"));
		try {
			// 3，根据需求请求所需数据
			String reportFilePath = "";
			List<String> detectUuidList = StringUtils.split(detectUuids, ",");
			String detectUuid = detectUuidList.get(0);
			Map<String, Object> detectInfoMap = new HashMap<String, Object>();
			detectInfoMap.put("detectUuid", detectUuid);
			List<Map<String, Object>> detectAndItemGvList = (List<Map<String, Object>>) detectInfoService.findItemAndDetectByDetectUuid(detectInfoMap).get("entityList");
			//4，封装数据到request中
			for (Map<String, Object> detectAndItemGv : detectAndItemGvList) {
				String resultNo = (String) detectAndItemGv.get("resultNo");
				String gfjgResultNo = (String) detectAndItemGv.get("gfjgResultNo");
				String sampleBarcode = (String) detectAndItemGv.get("sampleBarcode");
				request.setAttribute("statusUuid", detectAndItemGv.get("statusUuid"));
				if (OmsStatusConstant.DETECT_STATUS_UNEDIT.equals(detectAndItemGv.get("statusUuid")) ||
						OmsStatusConstant.DETECT_STATUS_UNTEST.equals(detectAndItemGv.get("statusUuid")) ||
						OmsStatusConstant.DETECT_STATUS_UNCHECK.equals(detectAndItemGv.get("statusUuid"))
						|| OmsStatusConstant.DETECT_STATUS_UNSIGN.equals(detectAndItemGv.get("statusUuid"))) {
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("isCertRevision")) && "Y".equals(detectAndItemGv.get("isCertRevision"))) {
						detectInfoMap.put("isCertRevision", "setNull");
						detectInfoService.updateByPrimaryKey(detectInfoMap);
					}
					// 获取客户英文名称。地址
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("custUuid"))) {
						Map<String, Object> custMap = new HashMap<String, Object>();
						custMap.put("custUuid", detectAndItemGv.get("custUuid"));
						Map<String, Object> customerInfoGv = (Map<String, Object>) (customerService.getCustomer(custMap).get("entityOne"));
						if (UtilValidate.isNotEmpty(customerInfoGv)) {
							if (UtilValidate.isNotEmpty(customerInfoGv.get("custNameEn"))) {
								request.setAttribute("custNameEn", customerInfoGv.get("custNameEn"));
							}
							if (UtilValidate.isNotEmpty(customerInfoGv.get("custAddrEn"))) {
								request.setAttribute("custAddrEn", customerInfoGv.get("custAddrEn"));
							}
						}
					}
					request.setAttribute("resultNo", resultNo);
					request.setAttribute("gfjgResultNo", gfjgResultNo);
					String[] detectStr = { "orderCode", "custName", "certCorpName", "sampleName", "manuFacturer", "custAddr", "certCorpAddr", "temData", "humData", "otherData", "labAddr",
							"certCorpCode","certPostCode","certLinkMan","certCorpTel","certCorpFax","labApprNo","resultConclus", "resultTypeUuId" ,"checkDate","orderType"};
					for (int i = 0; i < detectStr.length; i++) {
						if (UtilValidate.isNotEmpty(detectAndItemGv.get(detectStr[i]))) {
							request.setAttribute(detectStr[i], detectAndItemGv.get(detectStr[i]));
						} else {
							request.setAttribute(detectStr[i], "");
						}
					}
					if (detectUuidList.size() > 1) {
						request.setAttribute("specification", "见后");
						request.setAttribute("factoryCode", "见后");
					} else {
						request.setAttribute("specification", detectAndItemGv.get("specification"));
						request.setAttribute("factoryCode", detectAndItemGv.get("factoryCode"));
					}
					// 接收时间
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("checkDate"))) {
						//创建时间数据
						SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
						Calendar rightNow = Calendar.getInstance();
						Date currTestDate = sdf.parse(String.valueOf(detectAndItemGv.get("checkDate")));
						rightNow.setTime(currTestDate);
						//rightNow.add(Calendar.DAY_OF_YEAR, -1);//为检定时间前一天----违规
						Date date = rightNow.getTime();
						String greatDate = sdf.format(date);

						String[] checkDateTemp = greatDate.split("-");
						request.setAttribute("checkDateYear", checkDateTemp[0]);
						request.setAttribute("checkDateMonth", checkDateTemp[1]);
						request.setAttribute("checkDateDay", checkDateTemp[2]);
					}
					// 校准和测试类型的证书，结论插到证书后面的备注里
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("resultTypeUuId"))
							&& ("CONCLUS_TYPE_1".equals(detectAndItemGv.get("resultTypeUuId")) || "CONCLUS_TYPE_2".equals(detectAndItemGv.get("resultTypeUuId")))) {
						request.setAttribute("remark", detectAndItemGv.get("resultConclus"));
					}
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("currTestDate"))) {
						String[] currTestDateTemp = detectAndItemGv.get("currTestDate").toString().split("-");
						request.setAttribute("currTestDateYear", currTestDateTemp[0]);
						request.setAttribute("currTestDateMonth", currTestDateTemp[1]);
						request.setAttribute("currTestDateDay", currTestDateTemp[2]);
					}
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("endTestDate"))) {
						String[] sampleEndTestDateTemp = detectAndItemGv.get("endTestDate").toString().split("-");
						request.setAttribute("endTestDateYear", sampleEndTestDateTemp[0]);
						request.setAttribute("endTestDateMonth", sampleEndTestDateTemp[1]);
						request.setAttribute("endTestDateDay", sampleEndTestDateTemp[2]);
					}

					detectAndItemGv.put("orgUuid", userInfo.get("orgUuid"));// 设置当前所属机构
					detectAndItemGv.put("appSysUuid", userInfo.get("appSysUuid"));// 设置当前所属平台
					// 依据（规程）
					List<String> stdList = new ArrayList<String>();
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("stdUuid"))) {
						stdList = getStdList(detectAndItemGv);
					}
					request.setAttribute("stdList", stdList);
					// 标准装置
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("deviceUuid"))) {
						Map<String, Object> deviceMap = getDevice(detectAndItemGv);
						request.setAttribute("equipmentStandardList", deviceMap.get("equipmentStandardList"));
						request.setAttribute("accuracyLevelList", deviceMap.get("accuracyLevelList"));// 准确度等级
						request.setAttribute("measureRangeList", deviceMap.get("measureRangeList"));// 测量范围
						request.setAttribute("standardCount", deviceMap.get("standardCount"));// 数量
					}
					// 主标准器
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("stdEquipUuid")) && UtilValidate.isEmpty(detectAndItemGv.get("deviceUuid"))) {
						Map<String, Object> stdDeviceMap = getStdEquip(detectAndItemGv);
						request.setAttribute("equipmentList", stdDeviceMap.get("equipmentList"));
						request.setAttribute("measureRangeEquipList", stdDeviceMap.get("measureRangeEquipList"));
						request.setAttribute("accuracyLevelEquipList", stdDeviceMap.get("accuracyLevelEquipList"));
						request.setAttribute("equipCount", stdDeviceMap.get("equipCount") + "");

						request.setAttribute("equipmentStandardList", stdDeviceMap.get("equipmentStandardList"));
						request.setAttribute("accuracyLevelList", stdDeviceMap.get("accuracyLevelList"));// 不确定度
						request.setAttribute("measureRangeList", stdDeviceMap.get("measureRangeList"));// 测量范围
						request.setAttribute("standardCount", stdDeviceMap.get("standardCount") + "");
					}

				}
				// 根据检测单detectUuid得到检测结果
				Map<String, Object> resultInfoMap = new HashMap<String, Object>();
				resultInfoMap.put("detectUuid", detectUuid);
				Map<String, Object> resultInfoGv = (Map<String, Object>) resultInfoService.getResultByDetectUuid(resultInfoMap).get("entityOne");
				Map<String, Object> person = new HashMap<String, Object>();
				if (OmsStatusConstant.DETECT_STATUS_UNTEST.equals(detectAndItemGv.get("statusUuid"))) {
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("certTester"))) {
						String certTester = (String) detectAndItemGv.get("certTester");
						// 通过userUuid拿到这条person记录
						JSONObject personConfig = new JSONObject();
						personConfig.put("configUuid", "PersonSignature");
						personConfig.put("orgUuid", userInfo.get("orgUuid"));
						personConfig.put("appSysUuid", userInfo.get("appSysUuid"));
						
						//获取文件相关数据
						//获取文件配置信息
						Map<String, Object> configGv = fileConfigService.getFileConfig(personConfig);
						if (UtilValidate.isNotEmpty(configGv)) {
							Map<String,Object> configMap = (Map<String,Object>)configGv.get("entityOne");
							//获取文件的保存路径
							String personFolder = (String) configMap.get("fileFolder");
							JSONObject fileInfoJson = new JSONObject();
							fileInfoJson.put("entityName", "ump_user_person");
							fileInfoJson.put("entityRowId", certTester);
							//获取文件信息
							Map<String, Object> fileInfo = fileInfoService.getFileInfoByEntity(fileInfoJson);
							Map<String, Object> fileInfoEntityMap = (Map<String, Object>) fileInfo.get("entityOne");
							String filePath = "";
							//拼接文件路径
							if(UtilValidate.isNotEmpty(fileInfoEntityMap) && UtilValidate.isNotEmpty(fileInfoEntityMap.get("filePath"))) {
								filePath = (personFolder + fileInfoEntityMap.get("filePath"));
								File file = new File(filePath);
								if(file.exists())   {
									request.setAttribute("certTester_excel", filePath);
									filePath = filePath.replaceAll( "/","\\\\");
									request.setAttribute("certTester_word",filePath);
								}
							}
						}
					}
				}

				// 设置记录编号条形码
				JSONObject resultNoCodeConfigJson = new JSONObject();
				resultNoCodeConfigJson.put("configUuid", "ResultNoCodeDir");
				resultNoCodeConfigJson.put("orgUuid", userInfo.get("orgUuid"));
				resultNoCodeConfigJson.put("appSysUuid", userInfo.get("appSysUuid"));
				Map<String, Object> resultNoConfigEntity = fileConfigService.getFileConfig(resultNoCodeConfigJson);
				if(UtilValidate.isNotEmpty(resultNoConfigEntity)){
					Map<String,Object> twoMap = (Map<String,Object>)resultNoConfigEntity.get("entityOne");
					if(UtilValidate.isNotEmpty(resultNo)){
						String resultNoFileFolder = (String) twoMap.get("fileFolder");// 路径例如d:/lims_file/ResultNoCodeDir/
						//生成并创建二维码，返回文件路径
						String resultNoCodePath = createOneBarCode(resultNoFileFolder,resultNo,resultNo);
						String certBarCode = (String)detectAndItemGv.get("certBarCode"); 
						String certBarCodePath = createOneBarCode(resultNoFileFolder,certBarCode,certBarCode);
						//保存二维码路径到request中
						request.setAttribute("imagePath_zsbhtm_excel",certBarCodePath);
						resultNoCodePath = resultNoCodePath.replaceAll( "/","\\\\");
						request.setAttribute("imagePath_zsbhtm_word",resultNoCodePath);
					}
				}

				// 生成二维码
				Map<String, String> certMap = FastMap.newInstance();
				// 检定单位
				certMap.put("currTestDate", detectAndItemGv.get("currTestDate") + "");
				certMap.put("endTestDate", detectAndItemGv.get("endTestDate") + "");
				certMap.put("resultNo", detectAndItemGv.get("resultNo") + "");
				certMap.put("gfjgResultNo", detectAndItemGv.get("gfjgResultNo") + "");
				certMap.put("sampleName", detectAndItemGv.get("sampleName") + "");
				certMap.put("factoryCode", detectAndItemGv.get("factoryCode") + "");
				// 送检单位(证书单位)
				certMap.put("certCorpName", detectAndItemGv.get("certCorpName") + "");
				if (UtilValidate.isNotEmpty(detectAndItemGv.get("certTester"))) {
					// 通过userUuid拿到这条person记录
					JSONObject personJson = new JSONObject();
					personJson.put("userUuid", detectAndItemGv.get("certTester"));
					Map<String, Object> prsonList = userPersonService.getUserPerson(personJson);
					Map<String, Object> personMap = (Map<String, Object>) prsonList.get("entityOne");
					certMap.put("certTester", (String) personMap.get("firstName"));
					request.setAttribute("certTester",(String) personMap.get("firstName"));
				} else {
					certMap.put("certTester", "");

				}

				JSONObject twoCodeConfigJson = new JSONObject();
				twoCodeConfigJson.put("configUuid", "TwoBarCodeDir");
				twoCodeConfigJson.put("orgUuid", userInfo.get("orgUuid"));
				twoCodeConfigJson.put("appSysUuid", userInfo.get("appSysUuid"));
				Map<String, Object> twoCodeConfigEntity = fileConfigService.getFileConfig(twoCodeConfigJson);
				if(UtilValidate.isNotEmpty(twoCodeConfigEntity)){
					Map<String,Object> twoMap = (Map<String,Object>)twoCodeConfigEntity.get("entityOne");
					String twoFileFolder = (String) twoMap.get("fileFolder")+"/";// 路径例如d:/lims_file/JL_REPORT/
					String TwoBarCodePath = createTwoBarCode(twoFileFolder,"twoCode"+detectUuid,certMap);
					String twoJpg = TwoBarCodePath;
					request.setAttribute("imagePath_zsewm_excel",twoJpg);
					twoJpg = twoJpg.replaceAll( "/","\\\\");
					request.setAttribute("imagePath_zsewm_word",twoJpg);
				}
				if (OmsStatusConstant.DETECT_STATUS_UNTEST.equals(detectAndItemGv.get("statusUuid")) && "N".equals(detectAndItemGv.get("isCertUpload"))) {
					if (UtilValidate.isNotEmpty(resultInfoGv) && UtilValidate.isNotEmpty(resultInfoGv.get("resultUuid"))) {
						JSONObject fileInfoJson = new JSONObject();
						fileInfoJson.put("entityName", "ResultInfo");
						fileInfoJson.put("entityRowId", resultInfoGv.get("resultUuid"));
						fileInfoService.deleteFileInfo(fileInfoJson);
					}
				}
				// String reportFilePath = "d:\\test.doc";
				// word类型的
				if (UtilValidate.isNotEmpty(detectAndItemGv.get("fileTypeUuid")) && "WORD".equals(detectAndItemGv.get("fileTypeUuid"))) {
					mv = new ModelAndView("word_certReport_edit");
					// 报告文件地址

					// 原始记录
					String origFilePath = "";
					JSONObject fileInfoJson = new JSONObject();
					fileInfoJson.put("entityName", "ResultInfo");
					fileInfoJson.put("entityRowId", resultInfoGv.get("resultUuid"));
					fileInfoJson.put("mimeTypeId", "application/msword");
					Map<String, Object> fileInfo = fileInfoService.getFileInfoByEntity(fileInfoJson);
					Map<String, Object> fileInfoEntityMap = (Map<String, Object>) fileInfo.get("entityOne");
					//jsonParam.put("orgUuid", userInfo.get("orgUuid"));
					//jsonParam.put("appSysUuid", userInfo.get("appSysUuid"));

					if (UtilValidate.isNotEmpty(fileInfoEntityMap)) {
						reportFilePath = "downLoadFile?entityName=ResultInfo&entityRowId=" + resultInfoGv.get("resultUuid") +"&jsessionId="+jsessionId+"&mimeTypeId=application/msword";
					} else {
						if (UtilValidate.isNotEmpty(detectAndItemGv.get("resultTplUuid")) && "LastCertLogId".equals(detectAndItemGv.get("resultTplUuid"))) {
							reportFilePath = "downLoadFile?entityName=ResultInfo&entityRowId=" + detectAndItemGv.get("lastResultTplUuid") +"&jsessionId="+jsessionId+"&mimeTypeId=application/msword";
						} else {
							// 模板表中的数据
							reportFilePath = "downLoadFile?entityName=dms_tpl_cert&entityRowId=" + detectAndItemGv.get("resultTplUuid") +"&jsessionId="+jsessionId+"&mimeTypeId=application/msword";
							// 判断是否上传了原始记录模板
							JSONObject fileOrigJson = new JSONObject();
							fileOrigJson.put("entityName", "ResultInfoOrig");
							fileOrigJson.put("entityRowId", resultInfoGv.get("resultUuid"));
							Map<String, Object> fileInfoOrig = fileInfoService.getFileInfoByEntity(fileOrigJson);
							Map<String, Object> fileInfoOrigMap = (Map<String, Object>) fileInfoOrig.get("entityOne");
							// 判断是否上传了原始记录，拼接文件的下载路径
							if (UtilValidate.isNotEmpty(fileInfoOrigMap)) {
								origFilePath = "downLoadFile?entityName=ResultInfoOrig&entityRowId=" + resultInfoGv.get("resultUuid") +"&jsessionId="+jsessionId;
							} else {
								if (UtilValidate.isNotEmpty(detectAndItemGv.get("origTplUuid")) && "LastOrigLogId".equals(detectAndItemGv.get("origTplUuid"))) {
									origFilePath = "downLoadFile?entityName=TplOrig&entityRowId=" + detectAndItemGv.get("lastOrigTplUuid") +"&jsessionId="+jsessionId;
								}
							}
						}
					}
					request.setAttribute("reportFilePath", reportFilePath);
					request.setAttribute("origFilePath", origFilePath);
					// 判断文件类型，设置不同的jsp页面
				} else if (UtilValidate.isNotEmpty(detectAndItemGv.get("fileTypeUuid")) && "EXCEL".equals(detectAndItemGv.get("fileTypeUuid"))) {
					mv = new ModelAndView("word_certReport_excel_edit");
					if (OmsStatusConstant.DETECT_STATUS_UNTEST.equals(detectAndItemGv.get("statusUuid")) && UtilValidate.isEmpty(detectAndItemGv.get("isCertUpload"))) {
						if (UtilValidate.isNotEmpty(resultInfoGv) && UtilValidate.isNotEmpty(resultInfoGv.get("resultUuid"))) {
							JSONObject fileInfoJson = new JSONObject();
							fileInfoJson.put("entityName", "ResultInfo");
							fileInfoJson.put("entityRowId", resultInfoGv.get("resultUuid"));
							fileInfoJson.put("mimeTypeId", "application/msexcel");
							fileInfoService.deleteFileInfo(fileInfoJson);
						}
					}
					// 存放目录
					String excelFilePath = "";
					JSONObject fileConfigJson = new JSONObject();
					fileConfigJson.put("configUuid", "JL_REPORT");
					fileConfigJson.put("orgUuid", userInfo.get("orgUuid"));
					fileConfigJson.put("appSysUuid", userInfo.get("appSysUuid"));
					Map<String, Object> fileConfigEntity = (Map<String, Object>) fileConfigService.getFileConfig(fileConfigJson).get("entityOne");
					String fileFolder = (String) fileConfigEntity.get("fileFolder");// 路径例如d:/lims_file/JL_REPORT/
					// 报告文件地址
					Map<String, Object> fileInfoGv = new HashMap<String, Object>();
					String isFirstOpen = "";
					JSONObject fileInfoJson = new JSONObject();
					fileInfoJson.put("entityName", "ResultInfo");
					fileInfoJson.put("entityRowId", resultInfoGv.get("resultUuid"));
					fileInfoJson.put("mimeTypeId", "application/msexcel");
					Map<String, Object> fileInfo = fileInfoService.getFileInfoByEntity(fileInfoJson);
					Map<String, Object> fileInfoEntityMap = (Map<String, Object>) fileInfo.get("entityOne");
					if (UtilValidate.isNotEmpty(fileInfoEntityMap)) {
						reportFilePath = "downLoadFile?entityName=ResultInfo&entityRowId=" + resultInfoGv.get("resultUuid")+"&jsessionId="+jsessionId+"&mimeTypeId=application/msexcel";
						excelFilePath = fileFolder + fileInfoEntityMap.get("filePath").toString();
					} else {
						if (UtilValidate.isNotEmpty(detectAndItemGv.get("resultTplUuid")) && "LastResultId".equals(detectAndItemGv.get("resultTplUuid"))) {
							JSONObject fileInfoOldJson = new JSONObject();
							fileInfoOldJson.put("entityName", "ResultInfo");
							fileInfoOldJson.put("entityRowId", detectAndItemGv.get("lastResultTplUuid"));
							Map<String, Object> fileInfoOld = fileInfoService.getFileInfoByEntity(fileInfoOldJson);
							Map<String, Object> fileInfoEntityOldMap = (Map<String, Object>) fileInfoOld.get("entityOne");
							reportFilePath = "downLoadFile?entityName=ResultInfo&entityRowId=" + detectAndItemGv.get("lastResultTplUuid")+"&jsessionId="+jsessionId+"&mimeTypeId=application/msexcel";
							excelFilePath = fileFolder + fileInfoEntityOldMap.get("filePath");
							isFirstOpen = "Y";
						} else {
							JSONObject fileInfoTemplateJson = new JSONObject();
							fileInfoTemplateJson.put("entityName", "dms_tpl_cert");
							fileInfoTemplateJson.put("entityRowId", detectAndItemGv.get("resultTplUuid"));
							Map<String, Object> fileInfoTemplate = fileInfoService.getFileInfoByEntity(fileInfoTemplateJson);
							Map<String, Object> fileInfoEntityTemplateMap = (Map<String, Object>) fileInfoTemplate.get("entityOne");
							reportFilePath = "downLoadFile?entityName=dms_tpl_cert&entityRowId=" + detectAndItemGv.get("resultTplUuid")+"&jsessionId="+jsessionId+"&mimeTypeId=application/msexcel";
							JSONObject certTempConfigJson = new JSONObject();
							certTempConfigJson.put("configUuid", "CERT_TEMPLATE");
							certTempConfigJson.put("orgUuid", userInfo.get("orgUuid"));
							certTempConfigJson.put("appSysUuid", userInfo.get("appSysUuid"));
							Map<String, Object> certTempConfigEntity = (Map<String, Object>) fileConfigService.getFileConfig(certTempConfigJson).get("entityOne");
							String certFolder = (String) certTempConfigEntity.get("fileFolder");// 路径例如D:/uploadFile/tplcert
							excelFilePath = certFolder + StringUtils.getBackslash() +fileInfoEntityTemplateMap.get("filePath");
							isFirstOpen = "Y";
						}
					}
					// 填充报告文件
					if ((OmsStatusConstant.DETECT_STATUS_UNEDIT.equals(detectAndItemGv.get("statusUuid")))
							|| (OmsStatusConstant.DETECT_STATUS_UNTEST.equals(detectAndItemGv.get("statusUuid")) && UtilValidate.isNotEmpty(excelFilePath))) {
						String tempFilePath = getDefaultTaskOrderItemFilePath("JL_REPORT", detectUuid, "CERT",userInfo);
						String tempFileName = (String) detectAndItemGv.get("resultNo");
						String tempExtName = excelFilePath.substring(excelFilePath.lastIndexOf("."), excelFilePath.length());
						String outputFilePath = "";
						if (UtilValidate.isNotEmpty(isFirstOpen)) {
							outputFilePath = fileFolder + tempFilePath + StringUtils.getBackslash() + tempFileName + tempExtName;
						} else {
							outputFilePath = fileFolder + tempFilePath + StringUtils.getBackslash() + tempFileName + "Temp" + tempExtName;
						}
						String tempStatus = ExcelFileContent.setExcelFileContent(request, excelFilePath, outputFilePath, isFirstOpen, detectAndItemGv.get("resultTypeUuid").toString());
						if (tempStatus.equals("success")) {
							String myPath = tempFilePath + StringUtils.getBackslash() + tempFileName + tempExtName;
							// 保存记录fileInfo信息
							//判断文件表是否保存文件记录，否则不保存记录
							JSONObject fileJson = new JSONObject();
							fileJson.put("configUuid", "JL_REPORT");
							fileJson.put("entityName", "ResultInfo");
							fileJson.put("mimeTypeId", "application/msexcel");
							fileJson.put("entityRowId", resultInfoGv.get("resultUuid"));
							Map<String, Object> queryFileInfo = (Map<String, Object>) fileInfoService.getFileInfoByEntity(fileJson).get("entityOne");
							if (UtilValidate.isEmpty(queryFileInfo)) {
								fileJson.put("fileUuid", UtilUUID.uuidTomini());
								fileJson.put("fileName", tempFileName+tempExtName);
								fileJson.put("filePath", myPath);
								fileJson.put("mimeTypeId", "application/msexcel");
								fileJson.put("orgUuid", userInfo.get("orgUuid"));
								fileJson.put("appSysUuid", userInfo.get("appSysUuid"));
								fileInfoService.insertFileInfo(fileJson);
							}
							reportFilePath = "downLoadFile?entityName=ResultInfo&entityRowId=" + resultInfoGv.get("resultUuid")+"&jsessionId="+jsessionId+"&mimeTypeId=application/msexcel";
						}
					}
					request.setAttribute("resultTypeId", detectAndItemGv.get("resultTypeUuid").toString());
					request.setAttribute("statusUuid", detectAndItemGv.get("statusUuid"));
					request.setAttribute("reportFilePath", reportFilePath);
				}
				request.setAttribute("fileTypeId", detectAndItemGv.get("fileTypeUuid"));
				request.setAttribute("detectUuid", detectUuid);
				request.setAttribute("resultUuid",resultInfoGv.get("resultUuid"));
				//获取鉴定周期
				request.setAttribute("testDateCycle",detectAndItemGv.get("testDateCycle") + "");
			}

		} catch (Exception e) {
			String errMsg = "生成报告失败";
			request.setAttribute("_ERROR_MESSAGE_", errMsg);
		}
		// 证书参数
		// request.setAttribute("reportFilePath", "d:\\test.doc");
		request.setAttribute("poCtrl", poCtrl);
		// poCtrl.webOpen(path,OpenModeType.docAdmin,"中检远航");
		return mv;
	}

	/**
	 * 在线编辑原始记录
	 *
	 * @param request
	 * @return
	 */
	@RequestMapping(value = "/onlineEditOrig")
	public ModelAndView onlineEditOrig(HttpServletRequest request) {
		PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
		poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
		try {
			String detectUuids = request.getParameter("detectUuids");
			JSONObject jsonParams=new JSONObject();
			String jsessionId=request.getParameter("jsessionId");
			jsonParams.put("jsessionId", jsessionId);
			Map<String, Object> userInfo = getBaseDataToJSONObject(jsonParams);//获取到用户信息
			String userUuid = String.valueOf(userInfo.get("userUuid"));
			List<String> detectUuidList = StringUtils.split(detectUuids, ",");
			String detectUuid = detectUuidList.get(0);
			Map<String, Object> detectInfoMap = new HashMap<String, Object>();
			detectInfoMap.put("detectUuid", detectUuid);
			List<Map<String, Object>> detectAndItemGvList = (List<Map<String, Object>>) detectInfoService.findItemAndDetectByDetectUuid(detectInfoMap).get("entityList");
			Map<String, Object> detectAndItemGv = new HashMap<String, Object>();
			if (UtilValidate.isNotEmpty(detectAndItemGvList)) {
				detectAndItemGv = detectAndItemGvList.get(0);
				request.setAttribute("detectUuid", detectAndItemGv.get("detectUuid"));
				String resultNo = (String) detectAndItemGv.get("resultNo");
				String origLogNo = (String) detectAndItemGv.get("origLogNo");
				Map<String, Object> personMap = new HashMap<String, Object>();
				request.setAttribute("statusUuid", detectAndItemGv.get("statusUuid"));
				if (OmsStatusConstant.DETECT_STATUS_UNTEST.equals(detectAndItemGv.get("statusUuid"))) {
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("certTester"))) {
						String partyId = (String) detectAndItemGv.get("certTester");
						// 通过userUuid拿到这条person记录
						JSONObject personJson = new JSONObject();
						personJson.put("userUuid", partyId);
						Map<String, Object> prsonList = userPersonService.getUserPerson(personJson);
						personMap = (Map<String, Object>) prsonList.get("entityOne");
						if (UtilValidate.isNotEmpty(personMap.get("ufmFilePath"))) {
						}
					}
				} else if (OmsStatusConstant.DETECT_STATUS_UNCHECK.equals(detectAndItemGv.get("statusUuid"))) {
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("certCheckerUuid"))) {
						String certCheckerId = (String) detectAndItemGv.get("certCheckerUuid");
						// 通过userUuid拿到这条person记录
						JSONObject personJson = new JSONObject();
						personJson.put("userUuid", certCheckerId);
						Map<String, Object> prsonList = userPersonService.getUserPerson(personJson);
						personMap = (Map<String, Object>) prsonList.get("entityOne");
						if (UtilValidate.isNotEmpty(personMap.get("ufmFilePath"))) {
							// request.setAttribute("origLogChecker",
							// person.getString("ufmFilePath").replaceAll("\\\\","/"));

						}
					}
				}
				if (UtilValidate.isNotEmpty(personMap)) {
					request.setAttribute("firstName", personMap.get("firstName"));
				}
				if (OmsStatusConstant.DETECT_STATUS_UNTEST.equals(detectAndItemGv.get("statusUuid"))) {
					String taskOrderId = (String) detectAndItemGv.get("orderUuid");
					// 判断是否来自证书修改
					if (UtilValidate.isNotEmpty(detectAndItemGv) && UtilValidate.isNotEmpty(detectAndItemGv.get("isCertUpdated")) && "Y".equals(detectAndItemGv.get("isCertUpdated"))) {
						detectInfoMap.put("isCertRevision", "setNull");
						detectInfoService.updateByPrimaryKey(detectInfoMap);

					}
					String[] detectStr = { "resultNo", "gfjgResultNo", "origLogNo", "custName", "certCorpName", "custAddr", "sampleName", "specification", "manuFacturer", "temData", "humData",
							"otherData", "labAddr", "resultConclus", "resultTypeUuId", "factoryCode", "sampleBarcode" };
					for (int i = 0; i < detectStr.length; i++) {
						if (UtilValidate.isNotEmpty(detectAndItemGv.get(detectStr[i]))) {
							request.setAttribute(detectStr[i], detectAndItemGv.get(detectStr[i]));
						} else {
							request.setAttribute(detectStr[i], "");
						}
					}
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("currTestDate"))) {
						String[] currTestDateTemp = detectAndItemGv.get("currTestDate").toString().split("-");
						request.setAttribute("currTestDateYear", currTestDateTemp[0]);
						request.setAttribute("currTestDateMonth", currTestDateTemp[1]);
						request.setAttribute("currTestDateDay", currTestDateTemp[2]);
					}
					// 送样时间、接收时间
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("taskCheckDate"))) {
						String[] checkDateTemp = detectAndItemGv.get("checkDate").toString().split("-");
						request.setAttribute("checkDateYear", checkDateTemp[0]);
						request.setAttribute("checkDateMonth", checkDateTemp[1]);
						request.setAttribute("checkDateDay", checkDateTemp[2]);

					}
					// 依据（规程）
					List<String> stdList = new ArrayList<String>();
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("stdUuid"))) {
						stdList = getStdList(detectAndItemGv);
					}
					request.setAttribute("stdList", stdList);
					// 主标准器
					if (UtilValidate.isNotEmpty(detectAndItemGv.get("stdEquipUuid")) && UtilValidate.isEmpty(detectAndItemGv.get("deviceUuid"))) {
						Map<String, Object> stdDeviceMap = getStdEquip(detectAndItemGv);
						request.setAttribute("equipmentList", stdDeviceMap.get("equipmentList"));
						request.setAttribute("measureRangeEquipList", stdDeviceMap.get("measureRangeEquipList"));
						request.setAttribute("accuracyLevelEquipList", stdDeviceMap.get("accuracyLevelEquipList"));
						request.setAttribute("equipCount", stdDeviceMap.get("equipCount") + "");

						request.setAttribute("equipmentStandardList", stdDeviceMap.get("equipmentStandardList"));
						request.setAttribute("accuracyLevelList", stdDeviceMap.get("accuracyLevelList"));// 不确定度
						request.setAttribute("measureRangeList", stdDeviceMap.get("measureRangeList"));// 测量范围
						request.setAttribute("standardCount", stdDeviceMap.get("standardCount") + "");
					}
				}
				// 设置记录编号条形码
				if (UtilValidate.isNotEmpty(origLogNo)) {
					// String barCodeResultNoPath =
					// JlTaskHelper.getTaskOrderItemBarCodePath(origLogNo,"barcode.img.suffix");
					// request.setAttribute("imagePath_jlbhtm","/"+UtilHttp.getApplicationName(request)+barCodeResultNoPath.replaceAll("\\\\","/"));
				}
			}

			// 原始记录文件地址
			String origFilePath = "";
			String origFilePathParent = "";
			// 根据检测单detectUuid得到检测结果
			Map<String, Object> resultInfoMap = new HashMap<String, Object>();
			resultInfoMap.put("detectUuid", detectUuid);
			Map<String, Object> resultInfoGv = (Map<String, Object>) resultInfoService.getResultByDetectUuid(resultInfoMap).get("entityOne");
			JSONObject fileInfoJson = new JSONObject();
			fileInfoJson.put("entityName", "ResultInfoOrig");
			fileInfoJson.put("entityRowId", resultInfoGv.get("resultUuid"));
			Map<String, Object> fileInfo = fileInfoService.getFileInfoByEntity(fileInfoJson);
			Map<String, Object> fileInfoEntityMap = (Map<String, Object>) fileInfo.get("entityOne");
			if (UtilValidate.isNotEmpty(fileInfoEntityMap)) {
				origFilePath = "downloadCommonFile?entityName=ResultInfoOrig&entityRowId=" + resultInfoGv.get("resultUuid");
			} else {
				if (UtilValidate.isNotEmpty(detectAndItemGv.get("origTplUuid")) && "LastOrigLogId".equals(detectAndItemGv.get("origTplUuid"))) {
					origFilePath = "downloadCommonFile?entityName=ResultInfoOrig&entityRowId=" + detectAndItemGv.get("lastOrigTplUuid");
				} else {
					origFilePath = "downloadCommonFile?entityName=dms_tpl_orig&entityRowId=" + detectAndItemGv.get("origTplUuid");
					JSONObject jsonParam = new JSONObject();
					jsonParam.put("tplOrigUuid", detectAndItemGv.get("origTplUuid"));
					jsonParam.put("orgUuid", userInfo.get("orgUuid"));
					jsonParam.put("appSysUuid", userInfo.get("appSysUuid"));
					Map<String, Object> tplOrgMap = tplOrgService.getTplOrg(jsonParam);
					Map<String, Object> tplOrg = (Map<String, Object>) tplOrgMap.get("entityOne");// 得到所属封皮
					if (UtilValidate.isNotEmpty(tplOrg)) {
						origFilePathParent = "downloadCommonFile?entityName=dms_tpl_orig&entityRowId=" + tplOrg.get("origPurposeId");// 所属封皮的id
					}
				}
			}
			request.setAttribute("origFilePathParent", origFilePathParent);
			request.setAttribute("reportFilePath", origFilePath);
		} catch (Exception e) {
			String errMsg = "生成原始记录失败";
			request.setAttribute("_ERROR_MESSAGE_", errMsg);
		}
		// 原始记录参数
		request.setAttribute("poCtrl", poCtrl);
		ModelAndView mv = new ModelAndView("word_origReport_edit");
		return mv;
	}

	/**
	 * 查询对应主标准器与配套设备根据主键UUID
	 * @param map
	 * @return
	 */
	@SuppressWarnings("unchecked")
	private Map<String, Object> getStdEquip(Map<String, Object> map) {
		Map<String, Object> stdDeviceMap = new HashMap<String, Object>();
		try {
			List<Map<Object, Object>> equipmentStandardIdList = new ArrayList<Map<Object, Object>>();
			List<String> measureRangeEquipList = new ArrayList<String>();
			List<String> accuracyLevelEquipList = new ArrayList<String>();
			JSONObject equipJson = new JSONObject();
			equipJson.put("orgUuid", map.get("orgUuid"));
			equipJson.put("appSysUuid", map.get("appSysUuid"));
			equipJson.put("equipUuid", map.get("stdEquipUuid"));
			// 通过器具id找到所有的器具信息
			Map<String, Object> equipInfoMap = equipInfoService.findEquipBySample(equipJson);
			List<Map<String, Object>> euqipInfoList = (List<Map<String, Object>>) equipInfoMap.get("entityList");
			// 器具名称、规格型号、出厂编号相同的器具，证书编号/有效期整合在一起（用于解决一个标准器对应多个编号和有效期的问题）
			HashSet<Object> equipHashSet = new HashSet<Object>();
			Map<Object, Object> equipMap = FastMap.newInstance();
			int i = 0;
			String agencyName = "";
			for (Map<String, Object> euqipInfoGv : euqipInfoList) {
				/*if (i == 0 && UtilValidate.isNotEmpty(euqipInfoGv) && UtilValidate.isNotEmpty(euqipInfoGv.get("agencyUuid"))) {
					Map<String, Object> agencyUuidMap = new HashMap<String, Object>();
					agencyUuidMap.put("subcontractorUuid", euqipInfoGv.get("agencyUuid"));
					Map<String, Object> agencyMap = subcontracterService.getSubcontracter(agencyUuidMap);
					List<Map<String, Object>> agencyList = (List<Map<String, Object>>) agencyMap.get("entityList");
					if (UtilValidate.isNotEmpty(agencyList) && UtilValidate.isNotEmpty(agencyList.get(0).get("name"))) {
						agencyName = (String) agencyList.get(0).get("name");
					}
				}
				i++;
				stdDeviceMap.put("agencyName", agencyName);*/
				String equipType = euqipInfoGv.get("equipName").toString() + euqipInfoGv.get("equipModel").toString() + euqipInfoGv.get("factoryNo").toString();
				if (equipHashSet.contains(equipType)) {
					Map<Object, Object> eqTypeMap = (Map<Object, Object>) equipMap.get(equipType);
					String certNo = eqTypeMap.get("certNo") + "\n" + euqipInfoGv.get("certNo").toString();
					String measureRange = eqTypeMap.get("measureRange") + ";" + euqipInfoGv.get("measureRange");
					String accuracyLevel = (String) euqipInfoGv.get("accuracyLevel");
					eqTypeMap.put("certNo", certNo);
					measureRangeEquipList.set(measureRangeEquipList.indexOf(eqTypeMap.get("measureRange")), measureRange);
					accuracyLevelEquipList.set(accuracyLevelEquipList.indexOf(eqTypeMap.get("accuracyLevel")), accuracyLevel);
					eqTypeMap.put("measureRange", measureRange);
					eqTypeMap.put("accuracyLevel", accuracyLevel);
					eqTypeMap.put("vaildDate", euqipInfoGv.get("testVaildDate"));
				} else {
					equipHashSet.add(equipType);
					Map<Object, Object> equipmentStandardMap = FastMap.newInstance();
					if (UtilValidate.isNotEmpty(euqipInfoGv)) {
						String equipName = euqipInfoGv.get("equipName").toString();
						if (UtilValidate.isNotEmpty(euqipInfoGv.get("factoryNo"))) {
							equipmentStandardMap.put("factoryNo", euqipInfoGv.get("factoryNo"));
						}
						equipmentStandardMap.put("equipName", equipName);
						equipmentStandardMap.put("measureRange", euqipInfoGv.get("measureRange"));
						if (UtilValidate.isNotEmpty(euqipInfoGv.get("measureRange"))) {
							measureRangeEquipList.add(euqipInfoGv.get("measureRange") + "");
							equipmentStandardMap.put("measureRange", euqipInfoGv.get("measureRange") + "");
						} else {
							equipmentStandardMap.put("measureRange", "");
							measureRangeEquipList.add("");
						}
						equipmentStandardMap.put("equipModel", euqipInfoGv.get("equipModel"));
						String accuracyLevelString = "";
						if (UtilValidate.isNotEmpty(euqipInfoGv.get("accuracyLevel"))) {
							accuracyLevelString = euqipInfoGv.get("accuracyLevel") + "";
						} else {
							accuracyLevelString = euqipInfoGv.get("accuracyLevel") + "";
						}
						equipmentStandardMap.put("accuracyLevel", accuracyLevelString);
						accuracyLevelEquipList.add(accuracyLevelString);
						equipmentStandardMap.put("certNo", euqipInfoGv.get("certNo"));
						equipmentStandardMap.put("testValidDate", euqipInfoGv.get("testValidDate"));
						equipmentStandardMap.put("orderBy", euqipInfoGv.get("controlCode"));
						equipMap.put(equipType, equipmentStandardMap);
					}
				}
			}
			for (Map.Entry<Object, Object> entry : equipMap.entrySet()) {
				equipmentStandardIdList.add((Map<Object, Object>) entry.getValue());
			}
			stdDeviceMap.put("equipmentList", equipmentStandardIdList);
			stdDeviceMap.put("measureRangeEquipList", measureRangeEquipList);
			stdDeviceMap.put("accuracyLevelEquipList", accuracyLevelEquipList);
			stdDeviceMap.put("equipCount", equipmentStandardIdList.size() + "");
			stdDeviceMap.put("equipmentStandardList", equipmentStandardIdList);
			stdDeviceMap.put("accuracyLevelList", accuracyLevelEquipList);// 不确定度
			stdDeviceMap.put("measureRangeList", measureRangeEquipList);// 测量范围
			stdDeviceMap.put("standardCount", equipmentStandardIdList.size() + "");
		} catch (Exception e) {
			e.printStackTrace();
			String errorMessage = "主标准器相关信息失败";
			logger.error(errorMessage, e);
		}
		return stdDeviceMap;
	}

	/**
	 * 创建条码
	 * @param webappPath
	 * @param barCodeName
	 * @param barCodeString
	 * @throws IOException
	 */
	public static String createOneBarCode(String webappPath,String barCodeName,String barCodeString) throws IOException{
		BarCodeUtil.createBarCode(webappPath, barCodeName+".jpg", barCodeString);
		return webappPath + barCodeName + ".jpg";
	}

	/**
	 * 创建证书二维码
	 *
	 * @param webappPath
	 *            d:/twoCode
	 * @param barCodeName
	 * @param contentMap
	 * @throws IOException
	 */
	public static String createTwoBarCode(String webappPath, String barCodeName, Map<String, String> contentMap) throws IOException {
		barCode(webappPath, barCodeName + ".jpg", 194, 194, contentMap);
		return webappPath + barCodeName + ".jpg";
	}

	/**
	 * 二维码 标准版
	 *
	 * @param imgPath
	 * @param contentMap
	 */
	public static void barCode(String imgPath, String barCodeName, int width, int height, Map<String, String> contentMap) {
		try {
			if (!new File(imgPath).isDirectory()) {
				FileUtils.forceMkdir(new File(imgPath));
			}
			String content = "单位名称:中航工业北京长城计量测试技术研究所\n";
			if (UtilValidate.isNotEmpty(contentMap.get("gfjgResultNo"))) {
				content += "证书编号:" + contentMap.get("gfjgResultNo") + "\n";
			} else {
				content += "证书编号:" + contentMap.get("resultNo") + "\n";
			}
			if (UtilValidate.isNotEmpty(contentMap.get("certTester"))) {
				content += "检定员:" + contentMap.get("certTester") + "\n";
			} else {
				content += "检定员:\n";
			}
			if (UtilValidate.isNotEmpty(contentMap.get("sampleName"))) {
				content += "仪器名称:" + contentMap.get("sampleName") + "\n";
			} else {
				content += "仪器名称:\n";
			}
			if (UtilValidate.isNotEmpty(contentMap.get("factoryCode"))) {
				content += "出厂编号:" + contentMap.get("factoryCode") + "\n";
			} else {
				content += "出厂编号:\n";
			}
			content += "检定日期:" + contentMap.get("currTestDate") + "\n";
			content += "有效日期:" + contentMap.get("endTestDate") + "\n";
			content += "所属单位:" + contentMap.get("certCorpName") + "\n";
			MyZxingEncoderHandler handler = new MyZxingEncoderHandler();
			handler.encode(content, width, height, imgPath + "/" + barCodeName);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 以委托单的送检日期作为文件路径
	 * @param purposeId
	 *            文件路径ID
	 * @param detectUuid
	 *            检测单Id
	 * @param folderName
	 *            文件夹类型
	 * @return
	 * @throws Exception
	 */
	public String getDefaultTaskOrderItemFilePath(String purposeId, String detectUuid, String folderName,Map<String,Object> userInfo) throws Exception {
		JSONObject fileConfigJson = new JSONObject();
		fileConfigJson.put("configUuid", purposeId);
		// 根据检测单id查找样品条码信息
		Map<String, Object> detectInfoMap = new HashMap<String, Object>();
		detectInfoMap.put("detectUuid", detectUuid);
		List<Map<String, Object>> detectGvList = (List<Map<String, Object>>) detectInfoService.findItemAndDetectByDetectUuid(detectInfoMap).get("entityList");
		Map<String, Object> detectGv=detectGvList.get(0);
		String appendPath =  detectGv.get("checkDate").toString().replace("-","/")+ StringUtils.getBackslash() + detectGv.get("sampleBarcode")
				+ StringUtils.getBackslash() + folderName;
		fileConfigJson.put("appendPath", appendPath);
		fileConfigJson.put("orgUuid", userInfo.get("orgUuid"));
		fileConfigJson.put("appSysUuid", userInfo.get("appSysUuid"));
		// 根据configUuid和appendPath得到整个文件的路径信息
		Map<String, Object> fileConfigMap = (Map<String, Object>) fileConfigService.createFileDir(fileConfigJson).get("entityOne");
		String filePath = (String) fileConfigMap.get("filePath");// d:/lims_file/JL_REPORT/2018/09/27/样品条码/类型/
		if (UtilValidate.isNotEmpty(filePath)) {
			return appendPath;
		}
		return "";
	}

	/**
	 * 得到所有的规程的信息
	 * @return
	 */
	public List<String> getStdList(Map<String, Object> map) {
		List<String> stdList = new ArrayList<String>();
		try {
			String stName_code = "";
			for (String stdUuid : String.valueOf(map.get("stdUuid")).split(",")) {
				JSONObject stdJson = new JSONObject();
				stdJson.put("stdUuid", stdUuid);
				stdJson.put("orgUuid", map.get("orgUuid"));
				stdJson.put("appSysUuid", map.get("appSysUuid"));
				@SuppressWarnings("unchecked")
				Map<String, Object> stdGv = (Map<String, Object>) (standardService.getStd(stdJson)).get("entityOne");
				String stdCode = "";
				String stdName = "";
				if (UtilValidate.isNotEmpty(stdGv.get("stdCode"))) {
					stdCode = (String) stdGv.get("stdCode");
				}
				if (UtilValidate.isNotEmpty(stdGv.get("stdName"))) {
					stdName = (String) stdGv.get("stdName");
				}
				stName_code = stdCode + "《" + stdName + "》";
				stdList.add(stName_code);
			}
		} catch (Exception e) {
			e.printStackTrace();
			String errorMessage = "根据规程id得到规程集合";
			logger.error(errorMessage, e);
		}
		return stdList;
	}

	/**
	 * 得到所有的标准装置信息
	 * @return
	 */
	public Map<String, Object> getDevice(Map<String, Object> map) {
		Map<String, Object> deviceMap = new HashMap<String, Object>();
		try {
			List<Map<String, Object>> equipmentOrStandardList = new ArrayList<Map<String, Object>>();
			List<String> measureRangeList = new ArrayList<String>();
			List<String> accuracyLevelList = new ArrayList<String>();
			JSONObject deviceJson = new JSONObject();
			deviceJson.put("orgUuid", map.get("orgUuid"));
			deviceJson.put("appSysUuid", map.get("appSysUuid"));
			deviceJson.put("listEquipUuid", StringUtils.split(String.valueOf(map.get("deviceUuid")), ","));
			// 通过标准装置id找到所有的装置信息
			List<Map<String, Object>> equipInfoList = equipInfoService.findStdBySample(deviceJson);
			int standardCount = equipInfoList.size();
			for (Map<String, Object> equipInfoGv : equipInfoList) {
				Map<String, Object> equipDeviceMap = new HashMap<String, Object>();
				if (UtilValidate.isNotEmpty(equipInfoGv)) {
					equipDeviceMap.put("equipName", equipInfoGv.get("equipName"));
					equipDeviceMap.put("measureRange", equipInfoGv.get("measureRange"));
					if (UtilValidate.isNotEmpty(equipInfoGv.get("measureRange"))) {
						measureRangeList.add(equipInfoGv.get("measureRange") + "");
					} else {
						measureRangeList.add("");
					}
					equipDeviceMap.put("accuracyLevel", equipInfoGv.get("accuracyLevel"));
					accuracyLevelList.add(equipInfoGv.get("accuracyLevel") + "");
					equipDeviceMap.put("certNo", equipInfoGv.get("certNo"));
					equipDeviceMap.put("testValidDate", equipInfoGv.get("testValidDate"));
					equipDeviceMap.put("orderBy", equipInfoGv.get("equipUuid"));
				}
				equipmentOrStandardList.add(equipDeviceMap);
			}
			deviceMap.put("equipmentStandardList", equipmentOrStandardList);
			deviceMap.put("accuracyLevelList", accuracyLevelList);
			deviceMap.put("measureRangeList", measureRangeList);
			deviceMap.put("standardCount", standardCount);
		} catch (Exception e) {
			e.printStackTrace();
			String errorMessage = "根据装置id得到装置相关信息失败";
			logger.error(errorMessage, e);
		}
		return deviceMap;
	}

	/**
	 * 主标准器
	 * @return
	 */
	public Map<String, Object> getStdDevice(String stdEquipUuids) {
		Map<String, Object> stdDeviceMap = new HashMap<String, Object>();
		try {
			List<Map<Object, Object>> equipmentStandardIdList = new ArrayList<Map<Object, Object>>();
			List<String> measureRangeEquipList = new ArrayList<String>();
			List<String> accuracyLevelEquipList = new ArrayList<String>();
			JSONObject deviceJson = new JSONObject();
			deviceJson.put("listEquipUuid", StringUtils.split(stdEquipUuids, ","));
			// 通过器具id找到所有的器具信息
			Map<String, Object> equipInfoMap = equipInfoService.findEquipBySample(deviceJson);
			List<Map<String, Object>> euqipInfoList = (List<Map<String, Object>>) equipInfoMap;
			// 器具名称、规格型号、出厂编号相同的器具，证书编号/有效期整合在一起（用于解决一个标准器对应多个编号和有效期的问题）
			HashSet<Object> equipHashSet = new HashSet<Object>();
			Map<Object, Object> equipMap = FastMap.newInstance();
			int i = 0;
			String agencyName = "";
			for (Map<String, Object> euqipInfoGv : euqipInfoList) {
				if (i == 0 && UtilValidate.isNotEmpty(euqipInfoGv) && UtilValidate.isNotEmpty(euqipInfoGv.get("agencyUuid"))) {
					Map<String, Object> agencyUuidMap = new HashMap<String, Object>();
					agencyUuidMap.put("subcontractorUuid", euqipInfoGv.get("agencyUuid"));
					Map<String, Object> agencyMap = subcontracterService.getSubcontracter(agencyUuidMap);
					List<Map<String, Object>> agencyList = (List<Map<String, Object>>) agencyMap;
					if (UtilValidate.isNotEmpty(agencyList) && UtilValidate.isNotEmpty(agencyList.get(0).get("name"))) {
						agencyName = (String) agencyList.get(0).get("name");
					}
				}
				i++;
				stdDeviceMap.put("agencyName", agencyName);
				String equipType = euqipInfoGv.get("equipName").toString() + euqipInfoGv.get("equipModel").toString() + euqipInfoGv.get("factoryNo").toString();
				if (equipHashSet.contains(equipType)) {
					Map<Object, Object> map = (Map<Object, Object>) equipMap.get(equipType);
					String certNo = map.get("certNo") + "\n" + euqipInfoGv.get("certNo").toString();
					String measureRange = map.get("measureRange") + ";" + euqipInfoGv.get("measureRange");
					String accuracyLevel = (String) euqipInfoGv.get("accuracyLevel");
					map.put("certNo", certNo);
					measureRangeEquipList.set(measureRangeEquipList.indexOf(map.get("measureRange")), measureRange);
					accuracyLevelEquipList.set(accuracyLevelEquipList.indexOf(map.get("accuracyLevel")), accuracyLevel);
					map.put("measureRange", measureRange);
					map.put("accuracyLevel", accuracyLevel);
					map.put("vaildDate", euqipInfoGv.get("testVaildDate"));
				} else {
					equipHashSet.add(equipType);
					Map<Object, Object> equipmentStandardMap = FastMap.newInstance();
					if (UtilValidate.isNotEmpty(euqipInfoGv)) {
						String equipName = euqipInfoGv.get("equipName").toString();
						if (UtilValidate.isNotEmpty(euqipInfoGv.get("factoryNo"))) {
							equipmentStandardMap.put("factoryNo", euqipInfoGv.get("factoryNo"));
						}
						equipmentStandardMap.put("standardName", equipName);
						equipmentStandardMap.put("measureRange", euqipInfoGv.get("measureRange"));
						if (UtilValidate.isNotEmpty(euqipInfoGv.get("measureRange"))) {
							measureRangeEquipList.add(euqipInfoGv.get("measureRange") + "");
							equipmentStandardMap.put("measureRange", euqipInfoGv.get("measureRange") + "");
						} else {
							equipmentStandardMap.put("measureRange", "");
							measureRangeEquipList.add("");
						}
						equipmentStandardMap.put("equipModel", euqipInfoGv.get("equipModel"));
						String accuracyLevelString = "";
						if (UtilValidate.isNotEmpty(euqipInfoGv.get("accuracyLevel"))) {
							accuracyLevelString = euqipInfoGv.get("accuracyLevel") + "";
						} else {
							accuracyLevelString = euqipInfoGv.get("accuracyLevel") + "";
						}
						equipmentStandardMap.put("accuracyLevel", accuracyLevelString);
						accuracyLevelEquipList.add(accuracyLevelString);
						equipmentStandardMap.put("certNo", euqipInfoGv.get("certNo"));
						equipmentStandardMap.put("vaildDate", euqipInfoGv.get("testValidDate"));
						equipmentStandardMap.put("orderBy", euqipInfoGv.get("controlCode"));
						equipMap.put(equipType, equipmentStandardMap);
					}
				}
			}
			for (Map.Entry<Object, Object> entry : equipMap.entrySet()) {
				equipmentStandardIdList.add((Map<Object, Object>) entry.getValue());
			}
			stdDeviceMap.put("equipmentList", equipmentStandardIdList);
			stdDeviceMap.put("measureRangeEquipList", measureRangeEquipList);
			stdDeviceMap.put("accuracyLevelEquipList", accuracyLevelEquipList);
			stdDeviceMap.put("equipCount", equipmentStandardIdList.size() + "");

			stdDeviceMap.put("equipmentStandardList", equipmentStandardIdList);
			stdDeviceMap.put("accuracyLevelList", accuracyLevelEquipList);// 不确定度
			stdDeviceMap.put("measureRangeList", measureRangeEquipList);// 测量范围
			stdDeviceMap.put("standardCount", equipmentStandardIdList.size() + "");
		} catch (Exception e) {
			e.printStackTrace();
			String errorMessage = "主标准器相关信息失败";
			logger.error(errorMessage, e);
		}
		return stdDeviceMap;
	}

	/**
	 * 保存证书
	 *
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	@RequestMapping(value = "/saveFileCert")
	public void saveFileCert(HttpServletRequest request, HttpServletResponse response) throws Exception {
		FileSaver fs = new FileSaver(request, response);
		try {
			String jsessionId=request.getParameter("jsessionId");
			JSONObject jsonParam=new JSONObject();
			jsonParam.put("jsessionId", jsessionId);
			Map<String, Object> userInfo = getBaseDataToJSONObject(jsonParam);//获取到用户信息
			JSONObject json = new JSONObject();
			// 设置对应状态待检定、待核验、待签发
			json.put("statusTest", OmsStatusConstant.DETECT_STATUS_UNTEST);
			json.put("statusCheck", OmsStatusConstant.DETECT_STATUS_UNCHECK);
			json.put("statusSign", OmsStatusConstant.DETECT_STATUS_UNSIGN);
			String detectUuids = request.getParameter("detectUuids");// 获取到对应的检测单号
			json.put("detectUuids", detectUuids);
			json.put("userInfo", userInfo);
			json.put("checkDate",request.getParameter("checkDate"));//已委托单送检日期作为保存路径
			Map<String, Object> map = waitTestService.saveCert(json);
			String filePath = String.valueOf(map.get("filePath"));// 获取到对应文件夹
			String appendPath = String.valueOf(map.get("appendPath"));// 获取到相对路径
			String resultNo = String.valueOf(map.get("resultNo"));// 获取到对应证书号
			String fileExt = fs.getFileExtName();// 后缀
			String fileNameFull = resultNo+ fileExt; // 2018091000001.xlsx/2018091000001.doc
			fs.saveToFile(filePath + fileNameFull);// 保存证书模板文件
			String myPath = appendPath + StringUtils.getBackslash() + fileNameFull;
			JSONObject fileJson = new JSONObject();
			fileJson.put("configUuid", "JL_REPORT");
			fileJson.put("entityName", "ResultInfo");
			fileJson.put("entityRowId", map.get("resultUuid"));// 设置证书主键UUID
			fileJson.put("filePath", myPath);// 设置相对路径
			fileJson.put("fileName", fileNameFull);// 文件名称
			fileJson.put("fileSize", fs.getFileSize());// 文件大小
			fileJson.put("fileExtName", fileExt);// 文件后缀
			fileJson.put("fileDesc", resultNo);// 文件描述
			fileJson.put("comments", resultNo);// 文件描述
			fileJson.put("orgUuid", userInfo.get("orgUuid"));// 所属机构
			fileJson.put("appSysUuid", userInfo.get("appSysUuid"));// 所属平台

			/**
			 * 设置文件属性
			 */
			String fileType = "";
			if (UtilValidate.isNotEmpty(fileExt) && (".doc").equals(fileExt) || (".docx").equals(fileExt)) {// 说明是WROD
				fileType = "application/msword";
			} else if (UtilValidate.isNotEmpty(fileExt) && (".xls").equals(fileExt) || (".xlsx").equals(fileExt)) {
				fileType = "application/msexcel";
			} else {
				fileType = "OTHER";
			}
			fileJson.put("mimeTypeId", fileType);// 文件属性
			fileJson.put("createdTime", DateTimeUtils.getNowTimestamp());// 创建时间
			fileJson.put("createdUuid", userInfo.get("userUuid"));// 创建人
			//保存之前先删除之前记录,否则一个证书对应两条证书文件（因为为了在线打开excel文件在fileInfo表里存了一条记录）
			fileInfoService.deleteFileInfo(fileJson);
			// 保存文件信息
			fileInfoService.saveFileInfo(fileJson); // 保存文件信息
			// fs.saveToFile("d:\\lic\\" + fs.getFileName());//保存证书模板文件
			fs.setCustomSaveResult("ok");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (UtilValidate.isNotEmpty(fs)) {
				fs.close();
			}
		}
	}

	/**
	 * 保存委托单
	 *
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	@RequestMapping(value = "/saveOrderInfoFile")
	public void saveOrderInfoFile(HttpServletRequest request, HttpServletResponse response) throws Exception {
		FileSaver fs = new FileSaver(request, response);
		try {
			String orderCode=request.getParameter("orderCode");
			String orderUuid = request.getParameter("orderUuid");
			String checkDate = request.getParameter("checkDate");
			String jsessionId = request.getParameter("jsessionId");
			JSONObject jsonParam=new JSONObject();
			jsonParam.put("jsessionId", jsessionId);
			Map<String, Object> baseDataInfo = getBaseDataToJSONObject(jsonParam);//获取到用户信息
			// 上传报告文件
			// 根据configUuid查找委托单要存放的位置
			JSONObject fileConfigJson = new JSONObject();
			fileConfigJson.put("configUuid", "ORDER_SHEET_UPLOAD");
			fileConfigJson.put("orgUuid", baseDataInfo.get("orgUuid"));
			fileConfigJson.put("appSysUuid", baseDataInfo.get("appSysUuid"));
			Map<String, Object> fileConfigEntity = (Map<String, Object>) fileConfigService.getFileConfig(fileConfigJson).get("entityOne");
			String fileFolder = (String) fileConfigEntity.get("fileFolder");// 路径例如d:/lims_file/orderSheet/
			String appendPath = checkDate.replace("-","/")  + StringUtils.getBackslash();//拟取日期年+月+日
			fileConfigJson.put("appendPath", appendPath);
			//得到全路径
			String filePath = fileFolder + StringUtils.getBackslash() + appendPath ;// d:/lims_file/orderSheet/2018/09/27
			mkdirFilePath(filePath);//创建对应路径
			String fileExt = fs.getFileExtName();// 后缀
			String fileNameFull = orderCode + fileExt; // 2018091000001.xlsx/2018091000001.doc
			fs.saveToFile(filePath + fileNameFull);// 保存委托单文件
			String myPath = appendPath + StringUtils.getBackslash() + fileNameFull;
			JSONObject fileJson = new JSONObject();
			fileJson.put("configUuid", "ORDER_SHEET_UPLOAD");
			fileJson.put("entityName", "OrderInfo");
			fileJson.put("entityRowId", orderUuid);// 设置委托单主键UUID
			fileJson.put("filePath", myPath);// 设置相对路径
			fileJson.put("fileName", fileNameFull);// 文件名称
			fileJson.put("fileSize", fs.getFileSize());// 文件大小
			fileJson.put("fileExtName", fileExt);// 文件后缀
			fileJson.put("orgUuid", baseDataInfo.get("orgUuid"));// 所属机构
			fileJson.put("appSysUuid", baseDataInfo.get("appSysUuid"));// 所属平台
			/**
			 * 设置文件属性
			 */
			String fileType = "";
			if (UtilValidate.isNotEmpty(fileExt) && (".doc").equals(fileExt) || (".docx").equals(fileExt)) {// 说明是WROD
				fileType = "application/msword";
			} else if (UtilValidate.isNotEmpty(fileExt) && (".xls").equals(fileExt) || (".xlsx").equals(fileExt)) {
				fileType = "application/msexcel";
			} else {
				fileType = "OTHER";
			}
			fileJson.put("mimeTypeId", fileType);// 文件属性
			fileJson.put("createdTime", DateTimeUtils.getNowTimestamp());// 创建时间
			fileJson.put("createdUuid", baseDataInfo.get("userUuid"));// 创建人
			// 保存文件信息
			fileInfoService.saveFileInfo(fileJson); // 更新文件信息
			// fs.saveToFile("d:\\lic\\" + fs.getFileName());//保存证书模板文件
			fs.setCustomSaveResult("ok");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (UtilValidate.isNotEmpty(fs)) {
				fs.close();
			}
		}
	}




	/**
	 * 创建对应文件夹
	 * @param filePath
	 */
	private void mkdirFilePath(String filePath) {
		File file=new File(filePath);
		if(!file.exists()){//如果文件夹不存在
			file.mkdirs();//创建文件夹\支持层级目录
		}
	}


	@RequestMapping(value = "/index", method = RequestMethod.GET)
	public ModelAndView showIndex(HttpServletRequest request) {
		ModelAndView mv = new ModelAndView("index");
		return mv;
	}

	/**
	 * 打印样品条码
	 *
	 * @param request
	 * @return
	 */
	@RequestMapping(value = "/printBarcode")
	public ModelAndView printBarcode(HttpServletRequest request) {
		PDFCtrl poCtrl = new PDFCtrl(request);
		poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
		String uuids = request.getParameter("uuids");
		// 证书参数
		request.setAttribute("poCtrl", poCtrl);
		String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort();
		//设置文件的请求路径
		String pathUrl = basePath+"/fileInfo/downLoadFileByFilePath?uuids=" + uuids;
		//pageoffice请求路径并打开文件
		poCtrl.webOpen(pathUrl);
		ModelAndView mv = new ModelAndView("print_barcode");
		return mv;
	}

	/**
	 * 打印实验室条码
	 *
	 * @param request
	 * @return
	 */
	@RequestMapping(value = "/printLabBarcode")
	public ModelAndView printLabBarcode(HttpServletRequest request) {
		PDFCtrl poCtrl = new PDFCtrl(request);
		poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
		String uuids = request.getParameter("uuids");
		// 证书参数
		request.setAttribute("poCtrl", poCtrl);
		String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort();
		String pathUrl = basePath+"/fileInfo/downLoadLabFileByFilePath?uuids=" + uuids;
		poCtrl.webOpen(pathUrl);
		ModelAndView mv = new ModelAndView("print_barcode");
		return mv;
	}

	/**
	 * 打印委托单/外包交接单/交接单
	 *
	 * @param request
	 * @return
	 */
	@RequestMapping(value = "/printOrder")
	public ModelAndView printOrder(HttpServletRequest request) {
		PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
		ModelAndView mv = new ModelAndView();
		poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
		String orderUuid = request.getParameter("orderUuid");// 委托单id
		String jsessionId = request.getParameter("jsessionId");// 样品id
		String printType = request.getParameter("printType");// 打印类型
		String itemUuids = request.getParameter("itemUuids");//样品id
		String deptUuid = request.getParameter("deptUuid");//部门id
		JSONObject jsonParam = new JSONObject();
		jsonParam.put("orderUuid", orderUuid);
		jsonParam.put("jsessionId", jsessionId);
		request.setAttribute("orderUuid", orderUuid);
		request.setAttribute("jsessionId", jsessionId);
		try {
			if (UtilValidate.isNotEmpty(itemUuids)) {
				List<String> listItemUuid = StringUtils.split(itemUuids, ",");
				jsonParam.put("listItemUuid", listItemUuid);
			}
			if (UtilValidate.isNotEmpty(deptUuid)) {
				jsonParam.put("deptUuid", deptUuid);
			}
		/*
			//有检定deptUuid作为天剑，所以可以注释此段代码
			if (UtilValidate.isNotEmpty(printType) && printType.equals("printTaskOrderEx")) {// 打印外包交接单
				jsonParam.put("isSubcontract", "Y");
			} else if (UtilValidate.isNotEmpty(printType) && printType.equals("printTaskHandover")) {// 打印交接单
				jsonParam.put("isSubcontract", "N");
			}*/
			// 委托单和样品列表
			Map<String, Object> orderInfo = orderInfoController.getOrderInfo(jsonParam);
			Map<String, Object> entityOne = (Map<String, Object>) orderInfo.get(ResultMessageUtil.RESULT_ENTITY_ONE);
			//获取部门名称
			addDeptName(jsonParam, entityOne);
			/**获取状态名称*/
			addEnumName(jsonParam, orderInfo, "resultTypeUuid", "resultTypeUuid");
			addEnumName(jsonParam, entityOne, "resultTypeUuid", "resultTypeUuid");
			addEnumName(jsonParam, orderInfo, "payUuid", "payUuid");

			List<Map<String, Object>> entityList = (List<Map<String, Object>>) entityOne.get(ResultMessageUtil.RESULT_ENTITY_LIST);
			for (Map<String, Object> map : entityList) {
				if (UtilValidate.isNotEmpty(map.get("deptUuid"))) {
					Map<String, Object> subcontracterByUuid = subcontracterService.getSubcontracterByUuid(map.get("deptUuid").toString());
					if (UtilValidate.isNotEmpty(subcontracterByUuid)) {
						map.put("deptName", subcontracterByUuid.get("name").toString());
					}
				}
			}
			entityOne.put(ResultMessageUtil.RESULT_ENTITY_LIST, entityList);
			//样品标记为已打印
			List<String> listItemUuid = new ArrayList<>();
			for (Map<String, Object> map : entityList) {
				listItemUuid.add(map.get("itemUuid").toString());
			}
			jsonParam.put("listItemUuid", listItemUuid);
			// 证书参数
			jsonParam = getOrgAndAppSysInfo(jsonParam);// 得到当前登录人所有的信息
			entityOne.put("userName", jsonParam.get("userName"));
			entityOne.put("userUuid", jsonParam.get("userUuid"));
			Date date = (Date) jsonParam.get("updatedTime");
			SimpleDateFormat sdfs = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");// 时间格式
			String nowData = sdfs.format(date);
			entityOne.put("nowData", nowData);
			request.setAttribute("poCtrl", poCtrl);
			String entityName = "";
			String entityRowId = "";
			if (UtilValidate.isNotEmpty(printType) && printType.equals("printTaskOrder")) {// 打印委托单
				if (entityList.size() < 1) {
					request.setAttribute("EMPTY", "EMPTY");
				} else {
					entityName = "DEFAULT";
					entityRowId = "PRINT_WORDTASK";
				}
				mv = new ModelAndView("word_OrderPrint_edit_new");
			} /*else if (UtilValidate.isNotEmpty(printType) && printType.equals("printTaskOrderEx")) {// 打印外包交接单
				if (entityList.size() < 1) {
					request.setAttribute("EMPTY", "EMPTY");
				} else {
					entityName = "DEFAULT";
					entityRowId = "PRINT_WORDTASKOUT";
					jsonParam.put("isHandoverPrint", "Y");
					itemRequireService.updateItemRequire(jsonParam);
				}
				mv = new ModelAndView("word_OrderPrintOut_edit");
			}*/ else if (UtilValidate.isNotEmpty(printType) && printType.equals("printTaskHandover")) {// 打印交接单
				if (entityList.size() < 1) {
					request.setAttribute("EMPTY", "EMPTY");
				} else {
					entityOne.put("deptName", entityList.get(0).get("deptName"));
					jsonParam.put("isHandoverPrint", "Y");
					itemRequireService.updateItemRequire(jsonParam);
					entityName = "DEFAULT";
					entityRowId = "PRINT_TASKHANDOVER";
				}
				mv = new ModelAndView("word_TaskHandOver_edit");
			}
			request.setAttribute("entityOne", entityOne);
			request.setAttribute("entityList", entityList);
			request.setAttribute("sampleCount", entityList.size());
			String basePath = request.getScheme() + "://" + request.getServerName() + ":" + request.getServerPort();
			//默认模板
			String pathUrl =  basePath + "/fileInfo/downLoadFile?entityName=" + entityName + "&entityRowId=" + entityRowId + "&jsessionId=" + jsessionId;
			if (UtilValidate.isNotEmpty(entityName) && UtilValidate.isNotEmpty(entityRowId)) {
				if (UtilValidate.isNotEmpty(printType) && printType.equals("printTaskOrder")) {// 打印委托单
				//查询之前是否保存过委托单，有则调用
					JSONObject fileInfoJson = new JSONObject();
					fileInfoJson.put("entityName", "OrderInfo");
					fileInfoJson.put("entityRowId", entityOne.get("orderUuid"));
					Map<String, Object> fileInfo = fileInfoService.getFileInfoByEntity(fileInfoJson);
					Map<String, Object> fileInfoEntityMap = (Map<String, Object>) fileInfo.get("entityOne");
					if (UtilValidate.isNotEmpty(fileInfoEntityMap)) {
						pathUrl = basePath + "/fileInfo/downLoadFile?entityName=OrderInfo&entityRowId=" + entityOne.get("orderUuid") + "&jsessionId=" + jsessionId + "&mimeTypeId=application/msword";
					}
				}
			}
			poCtrl.webOpen(pathUrl, OpenModeType.docAdmin, "");
		}catch (Exception e) {
			String errMsg = "打印失败";
			request.setAttribute("_ERROR_MESSAGE_", errMsg);
		}
		return mv;
	}

	/**
	 * 证书打印
	 * @param request
	 * @return
	 * @throws Exception
	 */
	@RequestMapping(value = "/printCert")
	public ModelAndView printCert(HttpServletRequest request){
		String entityRowId = request.getParameter("entityRowId");
		String jsessionId = request.getParameter("jsessionId");
		String needPDF = request.getParameter("needPDF");
		String needPrint = request.getParameter("needPrint");
		String fileType = request.getParameter("fileType");
		String resultTypeId = request.getParameter("resultTypeId");
		ModelAndView modleView = new ModelAndView();
		//第一次请求时没有entityRowId，需要跳转到print_cert页面进行解析后，传相应的参数再跳转word_certReport_toPrint页面
		if (UtilValidate.isNotEmpty(entityRowId)) {
			request.setAttribute("entityRowId", entityRowId);
			request.setAttribute("jsessionId", jsessionId);
			request.setAttribute("needPDF", needPDF);
			request.setAttribute("needPrint", needPrint);
			request.setAttribute("fileType", fileType);
			request.setAttribute("resultTypeId", resultTypeId);
			modleView = new ModelAndView("word_certReport_toPrint");
		}else {
			modleView = new ModelAndView("print_cert");
		}

		return modleView;
	}

	/**
	 * 证书WORD转换一份PDF并保存
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	@RequestMapping(value = "/saveFileCertPdf")
	public void saveFileCertPdf(HttpServletRequest request, HttpServletResponse response) throws Exception {
		FileSaver fs = new FileSaver(request, response);
		try {
			String entityRowId = request.getParameter("entityRowId");
			String jsessionId=request.getParameter("jsessionId");
			JSONObject jsonParam=new JSONObject();
			jsonParam.put("jsessionId", jsessionId);
			jsonParam.put("entityRowId", entityRowId);
			jsonParam.put("resultUuid", entityRowId);
			Map<String, Object> userInfo = getBaseDataToJSONObject(jsonParam);//获取到用户信息
			String fileExt = "pdf";// 后缀
			Map<String, Object> findFileInfo = fileInfoService.findFileInfo(jsonParam);
			//获取上传的证书
			List<Map<String, Object>> entityList = (List<Map<String, Object>>) findFileInfo.get(ResultMessageUtil.RESULT_ENTITY_LIST);
			//判断pdf是否存在
			jsonParam.put("isPdf", "Y");
			Map<String, Object> resultByDetectUuid = resultInfoService.getResultByDetectUuid(jsonParam);
			Map<String, Object> entityOne  = (Map<String, Object>) resultByDetectUuid.get(ResultMessageUtil.RESULT_ENTITY_ONE);
			// 保存文件信息
			if (UtilValidate.isEmpty(entityOne) && UtilValidate.isNotEmpty(entityList) && entityList.size() > 0) {
				//设置证书打印是否为pdf为是
				Map<String, Object> dataDetect = new HashMap<String, Object>();
				dataDetect.put("isPdf", "Y");//是否是pdf
				dataDetect.put("resultUuid", entityRowId);//证书主键
				resultInfoService.updateResultInfo(dataDetect);
				//获取要更新的pdf文件数据
				Map<String, Object> map = new HashMap<String, Object>();
				for (Map<String, Object> hashMap : entityList) {
					String fileExtName = hashMap.get("fileExtName").toString();
					if (fileExtName.equals("pdf")) {
						map = hashMap;
						break;
					}
					map = hashMap;
				}
				//新增或更新文件数据
				JSONObject fileConfigJson = new JSONObject();
				fileConfigJson.put("configUuid", map.get("configUuid").toString());
				fileConfigJson.put("orgUuid", userInfo.get("orgUuid"));
				fileConfigJson.put("appSysUuid", userInfo.get("appSysUuid"));
				Map<String, Object> fileConfigEntity = fileConfigService.getFileConfig(fileConfigJson);
				Map<String, Object> fileConfig = (Map<String, Object>) fileConfigEntity.get(ResultMessageUtil.RESULT_ENTITY_ONE);
				String fileFolder = (String) fileConfig.get("fileFolder");
				map.put("fileExtName", fileExt);// 文件后缀
				String[] path = map.get("filePath").toString().split("\\.");
				String myPath = path[0]+"."+fileExt;
				map.put("filePath", myPath);// 设置相对路径
				myPath = fileFolder+myPath;
				fs.saveToFile(myPath);// 保存证书模板文件
				String[] fileName = map.get("fileName").toString().split("\\.");
				map.put("fileName", fileName[0]+"."+fileExt);// 文件名称
				map.put("mimeTypeId", "application/pdf");// 文件属性
				JSONObject jsonObject = JSONObject.parseObject(JSON.toJSONString(map));
				if (entityList.size() > 2) {
					jsonObject.put("updatedTime", DateTimeUtils.getNowTimestamp());// 创建时间
					jsonObject.put("updatedName", userInfo.get("updatedName"));// 创建人
					fileInfoService.updateFileInfo(jsonObject); // 更新文件信息
				}else {
					jsonObject.put("fileUuid",StringUtils.getUUID());
					jsonObject.put("createdTime", DateTimeUtils.getNowTimestamp());// 创建时间
					jsonObject.put("createdUuid", userInfo.get("createdUuid"));// 创建人
					fileInfoService.insertFileInfo(jsonObject); // 新增文件信息
				}
				fs.setCustomSaveResult("ok");
			}else {
				String errorMessage = "pdf证书已存在";
				logger.error(errorMessage);
			}
		} catch (Exception e) {
			e.printStackTrace();
			String errorMessage = "证书打印失败";
			logger.error(errorMessage, e);
		} finally {
			if (UtilValidate.isNotEmpty(fs)) {
				fs.close();
			}
		}
	}






}
