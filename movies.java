import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

// To build this program, you will need to add the following dependencies to your project:

// Apache POI: A Java library for reading and writing Microsoft Office file formats, including Excel spreadsheets.
// Apache POI-OOXML: A Java library for working with Office Open XML documents, which includes Excel spreadsheets.

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class MovieProgram {
    private static final String FILE_NAME = "movie_ratings.xlsx";
    private static final int MAX_ROWS = 100;

    private static final String[] COLUMN_HEADERS = {"Movie Name", "Rating", "Notes"};

    public static void main(String[] args) throws IOException {
        // Create a new Excel workbook and sheet
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Movie Ratings");

        // Create column headers
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < COLUMN_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(COLUMN_HEADERS[i]);
        }

        // Read existing data from the file, if it exists
        Map<String, Object[]> existingData = readExistingData();

        // Keep track of the number of rows we have written to the sheet
        int currentRow = 1;

        // Ask the user for movie ratings until we reach the maximum number of rows
        Scanner scanner = new Scanner(System.in);
        while (currentRow <= MAX_ROWS) {
            // Ask the user for the movie name, rating, and notes
            System.out.print("Enter the name of a movie: ");
            String movieName = scanner.nextLine();
            if (movieName.isEmpty()) {
                break; // stop asking for movies if the user presses enter without typing a name
            }

            System.out.print("Enter a rating for the movie (1-10): ");
            int rating = scanner.nextInt();
            scanner.nextLine(); // consume the newline character

            System.out.print("Enter any notes about the movie: ");
            String notes = scanner.nextLine();

            // Add the new movie rating to the existing data
            existingData.put(movieName, new Object[]{movieName, rating, notes});

            // Sort the movie ratings by rating, from most liked to most disliked
            List<Object[]> sortedData = new ArrayList<>(existingData.values());
            Collections.sort(sortedData, (a, b) -> (int)b[1] - (int)a[1]);

            // Write the data to the sheet
            for (int i = 0; i < sortedData.size(); i++) {
                if (i >= MAX_ROWS) {
                    break; // stop writing to the sheet if we have reached the maximum number of rows
                }

                Object[] rowValues = sortedData.get(i);
                Row row = sheet.createRow(currentRow++);
                for (int j = 0; j < rowValues.length; j++) {
                    Cell cell = row.createCell(j);
                    if (rowValues[j] instanceof String) {
                        cell.setCellValue((String)rowValues[j]);
                    } else if (rowValues[j] instanceof Integer) {
                        cell.setCellValue((int)rowValues[j]);
                    }
                }
            }
        }

        // Write the updated data to the file
        writeDataToFile(existingData, workbook);

        // Close the scanner and workbook
    }
}
