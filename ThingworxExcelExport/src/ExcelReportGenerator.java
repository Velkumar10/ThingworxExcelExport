import com.thingworx.entities.utils.ThingUtilities;
import com.thingworx.logging.LogUtilities;
import com.thingworx.metadata.annotations.ThingworxServiceDefinition;
import com.thingworx.metadata.annotations.ThingworxServiceParameter;
import com.thingworx.metadata.annotations.ThingworxServiceResult;
import com.thingworx.resources.Resource;
import com.thingworx.things.repository.FileRepositoryThing;
import com.thingworx.types.InfoTable;

import java.io.*;
import java.nio.file.Files;
import java.text.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.json.*;
import org.slf4j.Logger;

public class ExcelReportGenerator extends Resource {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private static Logger _logger = LogUtilities.getInstance().getApplicationLogger(ExcelReportGenerator.class);

	public ExcelReportGenerator() {
		// TODO Auto-generated constructor stub
	}

	@ThingworxServiceDefinition(name = "ExcelExport", description = "", category = "", isAllowOverride = false, aspects = {
			"isAsync:false" })
	@ThingworxServiceResult(name = "result", description = "Thingworx Download String\r\n", baseType = "STRING", aspects = {})
	public String ExcelExport(
			@ThingworxServiceParameter(name = "InfoTableData", description = "", baseType = "INFOTABLE", aspects = {
					"isRequired:true", "isEntityDataShape:true" }) InfoTable InfoTableData,
			@ThingworxServiceParameter(name = "RepositoryName", description = "", baseType = "THINGNAME", aspects = {
					"isRequired:true", "thingTemplate:FileRepository" }) String RepositoryName,
			@ThingworxServiceParameter(name = "ExcelTemplateLocation", description = "", baseType = "STRING") String ExcelTemplateLocation) throws Exception {
		_logger.trace("Entering Service: ExcelExport");
		// TODO Auto-generated method stub
				/******** STORE THE TEMPLATE LOCATION *****/
				String FILE_NAME = ExcelTemplateLocation;
				InfoTable a = InfoTableData; // Data to Infotable variable a
				Map<String, Integer> myMap = new HashMap<String, Integer>(); // Declare Map to store column header and key values
				/************ GET A COPY OF TEMPLATE ******/
				File fileToCopy = new File(FILE_NAME);
				String ext1 = getFileExtension(fileToCopy); // Call getFileExtension function to get File type
				DateFormat df = new SimpleDateFormat("ddMMyyyyHHmmss"); //Assign Date Format
				Date today = Calendar.getInstance().getTime(); //Get current Data and Time
				String reportDate = df.format(today); //Convert current Date and Time to Stirng Format as assigned
				String filename = null;
				// File[] drives = File.listRoots();
				// String location = drives[0].toString();
				FileRepositoryThing ExcelExporter;
				ExcelExporter = ((FileRepositoryThing)ThingUtilities.findThing("SystemRepository"));
				ExcelExporter.GetDirectoryStructure();
				filename = ExcelExporter.getRootPath() + File.separator + reportDate + "." + ext1;
				File newFile = new File(filename);
				Files.copy(fileToCopy.toPath(), newFile.toPath());
				/********** CREATE FILE ********/
				FileInputStream excelFile = new FileInputStream(new File(filename));
				if (ext1.equals("xlsx") || ext1.equals("xlsm")) {
					XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
					XSSFSheet sheet = workbook.getSheetAt(0);
					XSSFRow rows;
					JSONObject b = a.toJSON();
					JSONArray array = (JSONArray) b.get("rows");
					JSONObject datashapes = b.getJSONObject("dataShape");
					JSONObject fields = datashapes.getJSONObject("fieldDefinitions");
					int rowNum = 0;
					Iterator<?> keysvalue = fields.keys();
					rows = sheet.createRow(rowNum++);
					int colNum = 0;
					int keyseries = 0;
					while (keysvalue.hasNext()) {
						Cell cell = rows.createCell(colNum++);
						String keyvaluecolumn = (String) keysvalue.next();
						cell.setCellValue(keyvaluecolumn);
						myMap.put(keyvaluecolumn, keyseries++);
					}
					for (int i = 0; i < array.length(); i++) {
						JSONObject jsonObj2 = array.getJSONObject(i);
						Iterator<?> keysvalue2 = jsonObj2.keys();
						rows = sheet.createRow(rowNum++);
						colNum = 0;
						while (keysvalue2.hasNext()) {
							String Keyvaluesrow = (String) keysvalue2.next();
							int columnnum = myMap.get(Keyvaluesrow);
							Cell cell = rows.createCell(columnnum);
							if (jsonObj2.get(Keyvaluesrow) instanceof String) {
								if (jsonObj2.get(Keyvaluesrow) != null) {
									cell.setCellValue((String) jsonObj2.get(Keyvaluesrow));
								} else {
									cell.setCellValue("");
								}
							} else if (jsonObj2.get(Keyvaluesrow) instanceof Boolean) {
								if (jsonObj2.get(Keyvaluesrow) != null) {
									cell.setCellValue((Boolean) jsonObj2.get(Keyvaluesrow));
								} else {
									cell.setCellValue("");
								}
							} else if (jsonObj2.get(Keyvaluesrow) instanceof Integer) {
								if (jsonObj2.get(Keyvaluesrow) != null) {
									cell.setCellValue((Integer) jsonObj2.get(Keyvaluesrow));
								} else {
									cell.setCellValue((Integer) 0);
								}
							} else if (jsonObj2.get(Keyvaluesrow) instanceof Float) {
								if (jsonObj2.get(Keyvaluesrow) != null) {
									cell.setCellValue((Float) jsonObj2.get(Keyvaluesrow));
								} else {
									cell.setCellValue((Integer) 0);
								}
							} else if (jsonObj2.get(Keyvaluesrow) instanceof Long) {
								Date times = new Date((Long) jsonObj2.get(Keyvaluesrow));
								cell.setCellValue((Date) times);
							} else if (jsonObj2.get(Keyvaluesrow) instanceof Double) {
								if (jsonObj2.get(Keyvaluesrow) != null) {
									cell.setCellValue((Double) jsonObj2.get(Keyvaluesrow));
								} else {
									cell.setCellValue((Integer) 0);
								}
							}
						}
					}
					try {
						FileOutputStream outputStream = new FileOutputStream(filename);
						workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
						workbook.write(outputStream);
						workbook.close();

					} catch (FileNotFoundException e) {
						e.printStackTrace();
					} catch (IOException e) {
						e.printStackTrace();
					}
				} else if (ext1.equals("xls")) {
					HSSFWorkbook workbook = new HSSFWorkbook(excelFile);
					HSSFSheet sheet = workbook.getSheetAt(0);
					HSSFRow rows;
					JSONObject b = a.toJSON();
					JSONArray array = (JSONArray) b.get("rows");
					JSONObject datashapes = b.getJSONObject("dataShape");
					JSONObject fields = datashapes.getJSONObject("fieldDefinitions");
					int rowNum = 0;
					Iterator<?> keysvalue = fields.keys();
					rows = sheet.createRow(rowNum++);
					int colNum = 0;
					int keyseries = 0;
					while (keysvalue.hasNext()) {
						Cell cell = rows.createCell(colNum++);
						String keyvaluecolumn = (String) keysvalue.next();
						cell.setCellValue(keyvaluecolumn);
						myMap.put(keyvaluecolumn, keyseries++);
					}
					for (int i = 0; i < array.length(); i++) {
						JSONObject jsonObj2 = array.getJSONObject(i);
						Iterator<?> keysvalue2 = jsonObj2.keys();
						rows = sheet.createRow(rowNum++);
						colNum = 0;
						while (keysvalue2.hasNext()) {
							String Keyvaluesrow = (String) keysvalue2.next();
							int columnnum = myMap.get(Keyvaluesrow);
							Cell cell = rows.createCell(columnnum);
							if (jsonObj2.get(Keyvaluesrow) instanceof String) {
								if (jsonObj2.get(Keyvaluesrow) != null) {
									cell.setCellValue((String) jsonObj2.get(Keyvaluesrow));
								} else {
									cell.setCellValue("");
								}
							} else if (jsonObj2.get(Keyvaluesrow) instanceof Boolean) {
								if (jsonObj2.get(Keyvaluesrow) != null) {
									cell.setCellValue((Boolean) jsonObj2.get(Keyvaluesrow));
								} else {
									cell.setCellValue("");
								}
							} else if (jsonObj2.get(Keyvaluesrow) instanceof Integer) {
								if (jsonObj2.get(Keyvaluesrow) != null) {
									cell.setCellValue((Integer) jsonObj2.get(Keyvaluesrow));
								} else {
									cell.setCellValue((Integer) 0);
								}
							} else if (jsonObj2.get(Keyvaluesrow) instanceof Float) {
								if (jsonObj2.get(Keyvaluesrow) != null) {
									cell.setCellValue((Float) jsonObj2.get(Keyvaluesrow));
								} else {
									cell.setCellValue((Integer) 0);
								}
							} else if (jsonObj2.get(Keyvaluesrow) instanceof Long) {
								Date times = new Date((Long) jsonObj2.get(Keyvaluesrow));
								cell.setCellValue((Date) times);
							} else if (jsonObj2.get(Keyvaluesrow) instanceof Double) {
								if (jsonObj2.get(Keyvaluesrow) != null) {
									cell.setCellValue((Double) jsonObj2.get(Keyvaluesrow));
								} else {
									cell.setCellValue((Integer) 0);
								}
							}
						}
					}
					try {
						FileOutputStream outputStream = new FileOutputStream(filename);
						workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
						workbook.write(outputStream);
						workbook.close();

					} catch (FileNotFoundException e) {
						e.printStackTrace();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
				_logger.trace("Exiting Service: Exporter");
				String download = "/Thingworx/FileRepositories/SystemRepository/" + reportDate + "." + ext1;
				return download;
	}
	private String getFileExtension(File fileToCopy) {
		String fileName = fileToCopy.getName();
		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
			return fileName.substring(fileName.lastIndexOf(".") + 1);
		else
			return "";
	}

}
