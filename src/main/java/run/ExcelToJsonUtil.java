package run;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToJsonUtil {



	public static JSONObject excelBodyToJson(File FileName) throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException{


		Workbook Wb = getWorkbook(FileName.getAbsolutePath());

		int numbOfSheets = Wb.getNumberOfSheets(); // nombre de feuilles exel
		JSONObject Json = new JSONObject();

		for(int i=0;i<numbOfSheets;i++){

			Sheet sheet = Wb.getSheetAt(i); //

			boolean headingAvail = false; // on ne sait pas si il y a une entête
			int headingColnIndex = 0,headingRowIndex=0, lastRowIndex=0,lastCellIndex=0;
			//permet de formatter les données
			DataFormatter df = new DataFormatter();

			loop: for(Row ligne : sheet){ // on parcours toutes les lignes

				for(Cell cell : ligne){ // pour chaque ligne on parcours la colonne

					headingAvail = true; // on considère qu'il y a une entête
					headingColnIndex = cell.getColumnIndex();
					headingRowIndex = cell.getRowIndex();
					lastRowIndex = ligne.getLastCellNum();
					lastCellIndex = ligne.getLastCellNum();
					break loop;
				}
			}

			if(headingAvail){

				System.out.println("une entête est disponible");

				JSONArray JSheet = new JSONArray();

				for(int j= headingRowIndex+1;j<lastRowIndex;j++){

					JSONObject Jrow = new JSONObject();

					for(int k=headingColnIndex ; k<lastCellIndex;k++){
						Row Heading = sheet.getRow(headingRowIndex);
						Row row = sheet.getRow(j);
						Jrow.put(""+Heading.getCell(k), df.formatCellValue(row.getCell(k)));
					}

					JSheet.put(Jrow);
				}
				Json.put("Sheet "+i, JSheet);
			}
			else{
				System.out.println(" Heading is not available in the sheet"+(i+1));
				// traitement à faire si il n'y a pas d'entête
			}


		}

		System.out.print(Json.toString());
		return Json;
	}

	private static Workbook getWorkbook(String excelFilePath)
			throws IOException {
		Workbook workbook = null;

		if (excelFilePath.endsWith("xlsx")) {
			workbook = new XSSFWorkbook(excelFilePath);
		} else if (excelFilePath.endsWith("xls")) {
			workbook = new HSSFWorkbook(new FileInputStream(excelFilePath));
		} else {
			throw new IllegalArgumentException("The specified file is not Excel file");
		}

		return workbook;
	}

	private static void encodeFile(JSONObject excelContents, FileWriter file)
			throws IOException {
		try {
			file.write(excelContents.toString());
			System.out.println("succes encoding to json file");
		}
		catch (IOException e) {
			e.printStackTrace();
		}
		finally {
			file.flush();
			file.close();
		}
	}

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		String path = "src/main/resources/test.xlsx";
		File fileName = new File(path);
		//InputStream fileStream = ExcelToJson.class.getResourceAsStream("departement");
		JSONObject excelContents= excelBodyToJson(fileName);
		System.out.println(excelContents.toString());
		//FileWriter file = new FileWriter ("src/main/resources/"+fileName.getName()+".json");
		//encodeFile(excelContents, file);

	}


}

