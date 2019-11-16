
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.DirectoryNotEmptyException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Paths;
import java.sql.Connection;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

public class ExcelReader {
	private static String str;
	public static String SAMPLE_XLSX_FILE_PATH;
	private static String projectPath = System.getProperty("user.dir");
	private static String[] header={"Sum of Quantity","Sub_Inv"};
	private static List<String> headerModel = new ArrayList<String>();
	private static int modelCol = 0, subCol = 0, quantCol = 0,manufactureC = 0, descriptionC = 0;
	private static ArrayList<String> subList = new ArrayList<String>();
	private static Map<String, ArrayList<String[]>> map = new HashMap<String,ArrayList<String[]>>();
	public ExcelReader(File file) throws IOException, InvalidFormatException {
		this.SAMPLE_XLSX_FILE_PATH = projectPath+"\\ExcelFile\\input.xls";
		//Workbook workbook = null;
		FileInputStream inputStream = new FileInputStream(new File(projectPath+"\\ExcelFile\\input.xls"));
		Workbook workbook = null;
		try{
			if (SAMPLE_XLSX_FILE_PATH.endsWith("xlsx")) {
				workbook = new XSSFWorkbook(inputStream);
			} else if (SAMPLE_XLSX_FILE_PATH.endsWith("xls")) {
			//HSSFWorkbook
				workbook = new HSSFWorkbook(inputStream);
			}
		}catch(Exception e) {
			JOptionPane.showMessageDialog(null, "Cannot Read The File!");
		}
		generate(workbook,file);
		workbook.close();
		inputStream.close();
		// Closing the workbook
	}
	public void generate(Workbook workbook,File file) throws IOException {
		Sheet sheet = workbook.getSheetAt(0);
		//Write
		Workbook workbook1 = null;

		if (file.getName().endsWith(".xls")) {
			workbook1 = new HSSFWorkbook();
		}
		else if (file.getName().endsWith(".xlsx")){
			workbook1 = new XSSFWorkbook();
		}
		else {
			//System.out.println(file.getName());
			JOptionPane.showMessageDialog(null, "Invalid filename!");
			return;		
		}

		Sheet sheet1 = workbook1.createSheet("Details");

		DataFormatter dataFormatter = new DataFormatter();

		int count = 0;
		for (Row row: sheet) {
			if(count == 0) {
				int col = 0;
				for(Cell myCell: row) {
					String temp = dataFormatter.formatCellValue(myCell);
					if(temp.equals("Model")) {
						modelCol = col;
					}else if(temp.equals("Sub_Inv")) {
						subCol = col;
					}else if(temp.equals("Quantity")) {
						quantCol = col;
					}else if(temp.equals("Manufacturer")) {
						manufactureC = col;
					}else if(temp.equals("Description")) {
						descriptionC = col;
					}
					col++;
				}
				Row r = sheet1.createRow(count);
				createDataCell(r,(short) 0, "Sum of Quantity");
				createDataCell(r,(short) 1, "Sub_Inv");
				count++;
			}else {
				String model = row.getCell(modelCol).getStringCellValue();
				String sub = row.getCell(subCol).getStringCellValue();
				String manufacture = row.getCell(manufactureC).getStringCellValue();
				String description = row.getCell(descriptionC).getStringCellValue();
				int quantity = (int)row.getCell(quantCol).getNumericCellValue();
				if(!subList.contains(sub)) subList.add(sub);
				if(map.containsKey(model)) {
					//Map<String, ArrayList<String[]>>
					ArrayList<String[]> temp = map.get(model);
					boolean flag = false;
					for(int i = 0; i < temp.size(); i++) {
						if(sub.equals(temp.get(i)[0])) {
							int num = (int) Integer.parseInt(temp.get(i)[1]) +  quantity;
							temp.get(i)[1] = num+"";
							flag = true;
							break;
						}
					}
					if(!flag) {
						String[] tempa = {sub,quantity+"",manufacture,description};
						map.get(model).add(tempa);
					}
				}else{
					String[] tempa = {sub,quantity+"",manufacture,description};
					ArrayList<String[]> temp = new ArrayList<String[]>();
					temp.add(tempa);
					map.put(model, temp);
				}
			}
		}
		
		Row r = sheet1.createRow(count);
		int grandTotal  = subList.size();
		int[] sumA = new int[grandTotal+1];
		for(int i = 0; i < subList.size(); i++) {
			createDataCell(r,(short) (i+1), subList.get(i));
			if(i == subList.size()-1) {
				createDataCell(r,(short) (i+3), "Grand Total");
			}
		}
		
		count++;
		
		for (Map.Entry<String,ArrayList<String[]>> entry : map.entrySet()) {
			r = sheet1.createRow(count);
			String model = entry.getKey();
			ArrayList<String[]> list = entry.getValue();
			createDataCell(r,(short) 0, model);
			int sum = 0;
			for(String[] a : list) {
				int index = subList.indexOf(a[0])+1;
				sum+=Integer.parseInt(a[1]);
				sumA[index-1]+=Integer.parseInt(a[1]);
				createDataCell(r,(short) index, Integer.parseInt(a[1]));
			}
			sumA[grandTotal]+=sum;
			createDataCell(r,(short) (grandTotal+2), sum);
			count++;
		}
		r = sheet1.createRow(count+1);
		createDataCell(r,(short) 0, "Grand Total");
		for(int i = 0; i < grandTotal; i++) {
			createDataCell(r,(short) (i+1), sumA[i]);
		}
		createDataCell(r,(short) (grandTotal+2), sumA[grandTotal]);
		sheet1.setColumnWidth(0, 5000);
		sheet1.setColumnWidth(grandTotal+2, 5000);
		FileOutputStream outFile = new FileOutputStream(file);
		workbook1.write(outFile);
		outFile.flush();
		outFile.close();
		workbook1.close();
		workbook.close();
		deleteExcellFileFromClassPath();
	}
	public void deleteExcellFileFromClassPath() throws IOException {
		try
        { 
            Files.deleteIfExists(Paths.get(projectPath+"\\ExcelFile\\input.xls"));
            //Files.deleteIfExists(Paths.get(projectPath+"\\SumaryQuantity\\sumaryQuantity.xlsx"));
        } 
        catch(NoSuchFileException e) 
        { 
            System.out.println("No such file/directory exists"); 
        } 
        catch(DirectoryNotEmptyException e) 
        { 
            System.out.println("Directory is not empty."); 
        } 
        catch(IOException e) 
        { 
            System.out.println("Invalid permissions."); 
        }
	}
	
	private static void createHeaderCell(Row row, short col, String cellValue) {
		Cell c = row.createCell(col);
		c.setCellValue(cellValue);
	}
	
	private static void createDataCell(Row row, short col, String cellValue) {
		Cell c = row.createCell(col);
		c.setCellValue(cellValue);
	}

	private static void createDataCell(Row row, short col, int cellValue) {
		Cell c = row.createCell(col);
		c.setCellValue(cellValue);
	}
	
	public static int getModelQuantity() {
		return map.size();
	}

}