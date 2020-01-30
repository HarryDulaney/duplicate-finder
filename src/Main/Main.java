package Main;

import javafx.scene.Scene;
import javafx.scene.control.Label;

import java.io.File;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javafx.application.Application;
import javafx.geometry.Pos;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.VBox;
import javafx.scene.text.Font;
import javafx.stage.Stage;

public class Main extends Application {

	@Override
	public void start(Stage primaryStage) throws Exception {
		
		Font font = new Font(18);
		
		Label label = new Label("Welcome to DupChecker!\n"
				+ "\n"
				+ "This program takes in two columns of redundant data from different spreadsheets.xlsx\n"
				+ "and outputs a new Excel spreadsheet containing only unique values.\n" + "No Duplicates!!");
		
		label.setFont(font);
	
		label.setAlignment(Pos.BASELINE_CENTER);

		VBox vBox = new VBox(4);
		vBox.setAccessibleHelp("Click me to start.");
		Tooltip tp = new Tooltip("Click anywhere to start.");
		vBox.setAlignment(Pos.TOP_CENTER);

		Tooltip.install(vBox, tp);

		vBox.getChildren().add(label);

		Scene scene = new Scene(vBox,700,300);

		scene.setOnMousePressed(e -> {

			runDupCheck();

		});
		primaryStage.setScene(scene);

		primaryStage.show();

	}

	private void runDupCheck() {

		File fileOne = PHelper.showFilePrompt("Choose the file containing the first spreadsheet", ".xlsx");

		if (fileOne != null) {

			File fileTwo = PHelper.showFilePrompt(
					"Choose the file containing the second spreadsheet to compare to the first.", ".xlsx");
			if (fileTwo != null) {

				String colName = PHelper.showInputPrompt("Task data requested...",
						"Please enter the header for the column to check for duplicate values "
								+ "For Example: Email or Client ID Number: ",
						"Compare For Duplicates Task ");

				if (colName != null) {

					DuplicateChecker dupCheck = new DuplicateChecker(fileOne, fileTwo, colName);
					ArrayList<String> dupes = null;

					try {
						dupes = dupCheck.checkForDuplicates();

					} catch (Exception e) {
						e.printStackTrace();
					}

					Workbook workbook = new XSSFWorkbook();

					Sheet sheet = workbook.createSheet("Duplicates");

					Row headerRow = sheet.createRow(0);

					headerRow.createCell(0).setCellValue("Duplicate " + dupCheck.getColumnToCheck() + " 's");

					int j = 1;
					for (int i = 0; i < dupes.size(); i++) {

						Row row = sheet.createRow(j++);

						row.createCell(0).setCellValue(dupes.get(i));
					}

					sheet.autoSizeColumn(0);

					ExcelAccessObject.saveWorkbook(workbook);

				}
			}
		}

	}

	public static void main(String[] args) {
		launch(args);

	}

}
