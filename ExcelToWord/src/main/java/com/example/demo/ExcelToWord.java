package com.example.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.example.utils.XWPFUtils;
import com.monitorjbl.xlsx.StreamingReader;

public class ExcelToWord {

	private final static Logger logger = LoggerFactory.getLogger(ExcelToWord.class);
	List<String> queryColArray;// 要抓取的欄位
	List<String> tableOutputKey;// 要抓取的欄位
	String excelFolderPath; // excel資料夾位置
	String tempFileFolderPath; // word範本位置
	String destFileFolderPath; // word輸出資料夾位置
	String JCLNameLast;// 存放JCLName
	String systemName;
	String sheetName;
	XWPFUtils XWPFUtils = new XWPFUtils();

	// (CellIndex,HeaderName)
	Map<Integer, String> HeaderName = new HashMap<Integer, String>();

	ExcelToWord() {
		Properties pro = new Properties();
		// 設定檔位置
		String config = "config.properties";
		try {
			System.out.println("ExcelToWord 執行");
			// 讀取設定檔
			logger.info(MessageFormat.format("設定檔位置: {0}", config));
			pro.load(new FileInputStream(config));
			// 讀取資料夾位置
			excelFolderPath = pro.getProperty("excelDir");
			tempFileFolderPath = pro.getProperty("tempFile");
			destFileFolderPath = pro.getProperty("destFile");
			logger.info(" Excel資料夾位置:{}\n Word範例檔位置:{}\n 輸出資料夾位置:{}", excelFolderPath, tempFileFolderPath,
					destFileFolderPath);
			// 讀取需要抓取的欄位名稱
			queryColArray = Arrays.asList(pro.getProperty("queryColArray").split(","));
			tableOutputKey = Arrays.asList(pro.getProperty("tableOutputKey").split(","));
			logger.info("需要抓取的欄位 " + queryColArray);
		} catch (FileNotFoundException e) {
			logger.info(e.toString());
			e.printStackTrace();
		} catch (IOException e) {
			logger.info(e.toString());
			e.printStackTrace();
		} catch (Exception e) {
			logger.info(e.toString());
			e.printStackTrace();
		}
	}
	
	public void excelToWordStart() throws IOException {
		// 取資料夾
		File excelFolder = new File(excelFolderPath);
		logger.info("excelDir:{} 有 {} 個Excel檔案", excelFolderPath, excelFolder.list().length);
		for (File file : excelFolder.listFiles()) {
			logger.info("開始讀取 " + file.getName());
			// 讀取excel檔案
			Workbook wb = getExcelFile(file.getPath());
			if (wb == null) {
				logger.info(file.getPath() + "讀取失敗");
				throw new IOException("讀取失敗");
			}
			logger.info(file.getName() + " 讀取完成");
			// 解析Excel to List
			logger.info("開始解析 " + file.getName());
			List<Map<String, String>> excelInfoList = parseExcel(wb);
			logger.info("解析 " + file.getName() + " 完成");
			// excel檔名
			// EX: excel檔名 帳務作業流程清單(BANK)_1090430 取 帳務作業流程清單(BANK)
			systemName = file.getName();
			logger.info("預計輸出檔名 " + sheetName);
			outPutToWork(excelInfoList);
			logger.info(systemName + " 輸出完成");
		}
	}

	/**
	 * 讀取excel檔案
	 * 
	 * Workbook(Excel本體)、Sheet(內部頁面)、Row(頁面之行(橫的))、Cell(行內的元素)
	 * 
	 * 
	 * @param path excel檔案路徑
	 * @return excel內容
	 */
	public Workbook getExcelFile(String path) {
		Workbook wb = null;
		try {
			if (path == null) {
				return null;
			}
			String extString = path.substring(path.lastIndexOf(".")).toLowerCase();
			FileInputStream in = new FileInputStream(path);
			wb = StreamingReader.builder().rowCacheSize(100)// 存到記憶體行數，預設10行。
					.bufferSize(4096)// 讀取到記憶體的上限，預設1024
					.open(in);
		} catch (FileNotFoundException e) {
			logger.info(e.toString());
			e.printStackTrace();
		}

		return wb;
	}

	/**
	 * 解析Sheet
	 * 
	 * @param workbook Excel檔案
	 * @return 整個Sheet的資料
	 */
	public List<Map<String, String>> parseExcel(Workbook workbook) {
		// Sheet的資料
		List<Map<String, String>> excelDataList = new ArrayList<>();
		// 存放DNS欄位的欄位號
		int dnsIndex = 0;
		// 遍歷每一個sheet
		for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
			Sheet sheet = workbook.getSheetAt(sheetNum);
			boolean rowNum = true;

			sheetName = sheet.getSheetName();

			// 開始讀取sheet
			for (Row row : sheet) {
				// 先取header
				if (rowNum) {
					for (Cell cell : row) {
//						queryColArray.get(queryColArray.indexOf("JCL_Description"));
						if (queryColArray.contains(cell.getStringCellValue())) {
							if (cell.getStringCellValue() == "DSN") {
								dnsIndex = cell.getColumnIndex();
							} else {
								HeaderName.put(cell.getColumnIndex(), cell.getStringCellValue());
							}
						}
					}
					rowNum = false;
					continue;
				}
				/*
				 * OLD code Row firstRow = sheet.getRow(firstRowNum); if (null == firstRow) {
				 * System.out.println("解析Excel失敗"); } int rowStart = sheetNum;// 起始去掉首欄 int
				 * rowEnd = sheet.getPhysicalNumberOfRows();OLD int dnsIndex = 0; old for (Cell
				 * cell : firstRow) { if (cell.getStringCellValue().equals("DSN")) { dnsIndex =
				 * cell.getColumnIndex(); } } for (int rowNum = rowStart; rowNum < rowEnd;
				 * rowNum++) { Row row = sheet.getRow(rowNum); if (null == row) { continue; } //
				 * 解析Row的資料 excelDataList.add(convertRowToData(row, firstRow, dnsIndex)); }-
				 */
				// 解析Row的資料
				excelDataList.add(convertRowToData(row, dnsIndex));
			}
		}
		return excelDataList;
	}

	/**
	 * 將資料重組並輸出Word
	 * 
	 * @param excelDataList 整理過的Excel檔案
	 */
	public void outPutToWork(List<Map<String, String>> excelDataList) {
		// 抓出不重複的JCL
		HashSet<String> jclKeys = new HashSet<>();
		excelDataList.forEach(cn -> {
			jclKeys.add(cn.get("JCL"));
		});
		int fileCount = 0;
		logger.info("預計產出 {} 個檔案", jclKeys.size());
		// 將不重複的相同JCL_NAME的資料Group to List並輸出word
		for (String classKey : jclKeys) {
			logger.info("開始輸出:" + classKey);
			List<Map<String, String>> toWordList;
			toWordList = excelDataList.stream().filter(excelModel -> excelModel.get("JCL") == classKey)
					.collect(Collectors.toList()); // 篩選classkey之後回傳
			// 輸出Word
			createWord(toWordList);
			logger.info("已輸出 JCL Name: " + classKey);
			// 輸出完之後，刪除，節省資源。
			toWordList.forEach(Item -> excelDataList.remove(Item));
			toWordList.clear();
			fileCount++;
		}
		logger.info("實際產出 {} 個檔案", fileCount);
	}

	/**
	 * 解析ROW
	 * 
	 * @param row      資料行
	 * @param firstRow 標頭
	 * @param dnsIndex Dns的列數
	 * @return 整row的欄位
	 */
	public Map<String, String> convertRowToData(Row row, int dnsIndex) {
		Map<String, String> excelDateMap = new HashMap<String, String>();
		for (Cell cell : row) {
			// 1.先抓現在第幾個Column
			int cellNum = cell.getColumnIndex();
			// 2.再去抓Header的欄位名稱
			String firstRowName = HeaderName.get(cellNum);
			// 3.判斷是否為需要抓的欄位
			if (!queryColArray.contains(firstRowName) || firstRowName == null) {
				continue;
			}

			// "TWS AD Name,JCL Name,STEP Name,PROGRAM Name,DISP Status"
			// 抓到的欄位如果是JCL Name 會需要做空值塞值
			if (firstRowName.equals("JCL Name")) {
				if (cell.getStringCellValue().isEmpty() || cell.getStringCellValue() == null) {
					cell.setCellValue(JCLNameLast);
				} else {
					JCLNameLast = cell.getStringCellValue();
				}
			}
			// 如果是DISP Status，要抓DSN的值帶過來
			if (firstRowName.equals("DISP Status")) {
				switch (firstRowName) {
				case "MOD":
					firstRowName = "OUTPUT FILE";
					cell.setCellValue(row.getCell(dnsIndex).getStringCellValue());
					break;
				case "OLD":
					firstRowName = "INPUT FILE";
					cell.setCellValue(row.getCell(dnsIndex).getStringCellValue());
					break;
				case "SHR":
					firstRowName = "INPUT FILE";
					cell.setCellValue(row.getCell(dnsIndex).getStringCellValue());
					break;
				case "TLB645":
					break;
				default:
					break;
				}
			}

			excelDateMap.put(firstRowName, cell.getStringCellValue());
		}
		return excelDateMap;
	}

	public void createWord(List<Map<String, String>> excelList) {
		if (!new File(destFileFolderPath + "/" + systemName + "/" + sheetName).exists()) {
			new File(destFileFolderPath + "/" + systemName + "/" + sheetName).mkdirs();
		}
		try (InputStream is = new FileInputStream(tempFileFolderPath);
				OutputStream os = new FileOutputStream(destFileFolderPath + "/" + systemName + "/" + sheetName + "/"
						+ excelList.get(0).get("JCL") + ".docx");) {
			XWPFDocument doc = XWPFUtils.openDoc(is);
			List<XWPFParagraph> xwpfParas = doc.getParagraphs();

			List<Map<String, String>> Catalog = excelList.stream()
					.filter(item -> item.get("DISP_Initial").equals("SHR") || item.get("DISP_Initial").equals("OLD"))
					.collect(Collectors.toList());

			List<Map<String, String>> inputList = excelList
					.stream().filter(item -> item.get("OPEN_Mode").equals("INPUT")
							|| item.get("OPEN_Mode").equals("I-O") || item.get("OPEN_Mode").equals("INPUT,OUTPUT"))
					.collect(Collectors.toList());

			List<Map<String, String>> ouputList = excelList.stream()
					.filter(item -> item.get("OPEN_Mode").equals("OUTPUT") || item.get("OPEN_Mode").equals("I-O")
							|| item.get("OPEN_Mode").equals("INPUT,OUTPUT")
							|| item.get("OPEN_Mode").equals("I-O,INPUT"))
					.collect(Collectors.toList());

			for (XWPFParagraph xwpfParagraph : xwpfParas) {
				String itemText = xwpfParagraph.getText();
				switch (itemText) {
				case "${catalogTable}":
					XWPFUtils.replaceTable(doc, itemText, Catalog, tableOutputKey);
					break;
					
				case "${dataTable}":
					XWPFUtils.replaceTable(doc, itemText, inputList, tableOutputKey);
					break;

				case "${resultTable}":
					XWPFUtils.replaceTable(doc, itemText, ouputList, tableOutputKey);
					break;
				}
			}

			Map<String, Object> data = new HashMap<>();
			data.put("${測試案例_L2}", excelList.get(0).get("測試案例_L2"));
			data.put("${系統別}", excelList.get(0).get("系統別"));
			data.put("${AD_NAME}", excelList.get(0).get("AD"));
			data.put("${AD_Description}", excelList.get(0).get("AD_Description"));
			data.put("${JCL_NAME}", excelList.get(0).get("JCL"));
			data.put("${JCL_Description}", excelList.get(0).get("JCL_Description"));
//			data.put("${description}", "暫定");
			// 取代資料
			XWPFUtils.replaceInPara(doc, data);
			doc.write(os);
		} catch (FileNotFoundException e) {
			logger.info(e.toString());
			e.printStackTrace();
		} catch (IOException e) {
			logger.info(e.toString());
			e.printStackTrace();
		}
	}

}
