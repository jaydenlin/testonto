package testonto;

import asus.com.onto.ONTO;

import java.io.IOException;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

import javax.print.attribute.Size2DSyntax;

import org.apache.poi.ss.formula.CollaboratingWorkbooksEnvironment;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class testonto {

	public static void main(String[] args) {

		ExcelToOnto(args[0],1,args[2]);

	}
	
	static void ExcelToOnto(String excelFilePath,int sheetIndex,String outputRDFFileName){
		// set up excel
				String filePath = excelFilePath;//set path
				XSSFWorkbook xwb = null;
				try {
					xwb = new XSSFWorkbook(filePath);
				} catch (IOException e) {
					System.out.println("讀取excel失敗");
					e.printStackTrace();
				}

				// read data form excel
				XSSFSheet sheet = xwb.getSheetAt(sheetIndex);//get sheet
				// Domain
				String domain = sheet.getRow(0).getCell(1).getStringCellValue();
				// Topic
				String topic = sheet.getRow(1).getCell(1).getStringCellValue();
				
				///////////////
				// write into ontology
				//////////////
				
				String URI = "http://asus.com/onto/";
				String fileName = outputRDFFileName;//set file name
				ONTO onto = new ONTO();
				onto.createRDF(URI);

				// setup class
				List<String> entityClass1 = readAndSetClass(onto, sheet, 1);
				List<String> relationClass = readClassRelation(onto, sheet);
				List<String> entityClass2 = readAndSetClass(onto, sheet, 3);

				// setup instance
				List<String> entityInstance1 = readAndSetInstance(onto, sheet, 1,
						entityClass1.get(entityClass1.size() - 1));
				List<String> relationInstance = readInstanceRelation(onto, sheet);
				List<String> entityInstance2 = readAndSetInstance(onto, sheet, 3,
						entityClass2.get(entityClass2.size() - 1));

				// set class relation
				setRelation(onto, entityClass1, relationClass, entityClass2, true);
				
				// set instance relation
				setRelation(onto, entityInstance1, relationInstance, entityInstance2,
						false);

				onto.writeAndClose(fileName);
				onto.closeRDF();
	}

	static List<String> readAndSetClass(ONTO onto, XSSFSheet sheet,
			int columnIndex) {
		List<String> classList = new ArrayList<String>();
		String parentClass = sheet.getRow(3).getCell(columnIndex)
				.getStringCellValue().trim().split("\\(")[0].trim();
		String childClass;

		for (int row = 4; row < sheet.getPhysicalNumberOfRows(); row++) {
			if (sheet.getRow(row).getCell(columnIndex) == null) {
				break;
			}
			if (sheet.getRow(row).getCell(columnIndex).getStringCellValue()
					.equals("Entity Instance")) {
				break;
			}
			childClass = sheet.getRow(row).getCell(columnIndex)
					.getStringCellValue().trim().split("\\(")[0].trim();

			System.out.print(parentClass);
			System.out.print(" hasSubClass ");
			System.out.println(childClass);

			onto.addClassSubClass(parentClass, childClass);
			parentClass = childClass;
			classList.add(parentClass);
		}

		return classList;

	}

	static List<String> readAndSetInstance(ONTO onto, XSSFSheet sheet,
			int columnIndex, String className) {

		List<String> instanceList = new ArrayList<>();
		int instanceStartRow = 0;
		for (int row = 4; row < sheet.getPhysicalNumberOfRows(); row++) {
			if (sheet.getRow(row).getCell(1) == null) {
				break;
			}
			if (sheet.getRow(row).getCell(1).getStringCellValue()
					.equals("Entity Instance")) {
				instanceStartRow = row + 1;
				break;
			}
		}

		for (; instanceStartRow < sheet.getPhysicalNumberOfRows(); instanceStartRow++) {

			if (sheet.getRow(instanceStartRow).getCell(columnIndex) == null) {
				break;
			}
			String instanceName = sheet.getRow(instanceStartRow)
					.getCell(columnIndex).getStringCellValue().trim();

			System.out.print(className);
			System.out.print(" hasInstance ");
			System.out.println(instanceName);

			onto.addClassInstance(className, instanceName);
			instanceList.add(instanceName);
		}

		return instanceList;

	}

	static List<String> readClassRelation(ONTO onto, XSSFSheet sheet) {

		List<String> classList = new ArrayList<String>();
		int columnIndex = 2;
		String parentClass = sheet.getRow(3).getCell(columnIndex)
				.getStringCellValue().trim().split("\\(")[0].trim();
		String childClass;

		for (int row = 4; row < sheet.getPhysicalNumberOfRows(); row++) {
			if (sheet.getRow(row).getCell(columnIndex) == null) {
				break;
			}
			if (sheet.getRow(row).getCell(columnIndex).getStringCellValue()
					.equals("Relation Instance")) {
				break;
			}
			childClass = sheet.getRow(row).getCell(columnIndex)
					.getStringCellValue().trim().split("\\(")[0].trim();

			// onto.addClassSubClass(parentClass, childClass);
			parentClass = childClass;
			classList.add(parentClass);
		}

		return classList;

	}

	static List<String> readInstanceRelation(ONTO onto, XSSFSheet sheet) {

		List<String> instanceList = new ArrayList<>();
		int columnIndex = 2;
		int instanceStartRow = 0;
		for (int row = 4; row < sheet.getPhysicalNumberOfRows(); row++) {
			if (sheet.getRow(row).getCell(1) == null) {
				break;
			}
			if (sheet.getRow(row).getCell(1).getStringCellValue()
					.equals("Entity Instance")) {
				instanceStartRow = row + 1;
				break;
			}
		}

		for (; instanceStartRow < sheet.getPhysicalNumberOfRows(); instanceStartRow++) {

			if (sheet.getRow(instanceStartRow).getCell(columnIndex) == null) {
				break;
			}
			String instanceName = sheet.getRow(instanceStartRow)
					.getCell(columnIndex).getStringCellValue().trim();

			// System.out.print(className);
			// System.out.print(" hasInstance ");
			// System.out.println(instanceName);
			//
			// onto.addClassInstance(className, instanceName);
			instanceList.add(instanceName);
		}

		return instanceList;

	}

	static void setRelation(ONTO onto, List<String> entity1,
			List<String> relation, List<String> entity2, boolean isClass) {
		
		int length=Math.max(Math.max(entity1.size(),entity2.size()),relation.size());
		String entity1Temp="";
		String relationTemp="";
		String entity2Temp="";
		
		if (isClass) {
			for (int i = 0; i < length; i++) {

				if (i < entity1.size()) {
					entity1Temp=entity1.get(i);
				}

				if (i < relation.size()) {
					relationTemp=relation.get(i);
				}
				

				if (i < entity2.size()) {
					entity2Temp=entity2.get(i);
				}	
				
				onto.addClassRelation(entity1Temp,relationTemp,entity2Temp);
			}
		} else {
			for (int i = 0; i < length; i++) {
				
				if (i < entity1.size()) {
					entity1Temp=entity1.get(i);
				}

				if (i < relation.size()) {
					relationTemp=relation.get(i);
				}
				

				if (i < entity2.size()) {
					entity2Temp=entity2.get(i);
				}	
				
				onto.addInstanceRelation(entity1Temp,relationTemp,entity2Temp);
			}
		}
	}

}
