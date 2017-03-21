package run;
import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelToJsonUtil {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		String path = "src/main/resources/test.xlsx";
		File fileName = new File(path);
		//InputStream fileStream = ExcelToJson.class.getResourceAsStream("departement");
		JSONObject excelContents= excelBodyToJson(fileName);
		System.out.println(excelContents.toString());
	}

	public static JSONObject excelBodyToJson(File FileName) throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException{
		Workbook Wb = WorkbookFactory.create(FileName);
		int NoOfSheets = Wb.getNumberOfSheets();
		JSONObject Json = new JSONObject();

		for(int i=0;i<NoOfSheets;i++){
			Sheet sheet = Wb.getSheetAt(i);
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
			}else{
				System.out.println(" Heading is not available in the sheet"+(i+1));
			}


		}

		System.out.print(Json.toString());
		return Json;
	}

}

