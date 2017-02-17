import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.monitorjbl.xlsx.StreamingReader;

public class Main {

	public static void main(String[] args) throws IOException {
		File file = new File("excel.xlsx");

		Main main = new Main();
		main.runPOI(file);
		main.runStreamingReader(file);
	}

	public void runStreamingReader( File file )
	{
		FileInputStream inputStream = null;
		Workbook workbook = null;
		try {
			inputStream = new FileInputStream( file );
			workbook = StreamingReader.builder()
					.rowCacheSize( 100 )
					.bufferSize( 4096 )
					.open( inputStream );

			for ( Sheet sheet : workbook ) {
				System.out.println( sheet.getSheetName() );
				for ( Row r : sheet ) {
					for ( Cell c : r ) {
						System.out.println( c.getStringCellValue() );
					}
				}
			}
		}
		catch ( FileNotFoundException e ) {
		} finally {
			IOUtils.closeQuietly( workbook );
			IOUtils.closeQuietly( inputStream );
		}
	}

	public void runPOI( File file )
	{
		FileInputStream inputStream = null;
		Workbook workbook = null;

		try {
			inputStream = new FileInputStream( file );

			workbook = new XSSFWorkbook( inputStream );
			Iterator<Sheet> sheetIterator = workbook.sheetIterator();
			while ( sheetIterator.hasNext() ) {

				Sheet firstSheet = sheetIterator.next();
				Iterator<Row> iterator = firstSheet.rowIterator();
				while ( iterator.hasNext() ) {
					Row nextRow = iterator.next();

					Iterator<Cell> cellIterator = nextRow.cellIterator();
					while ( cellIterator.hasNext() ) {
						Cell cell = cellIterator.next();

						switch ( cell.getCellTypeEnum() ) {
						case STRING:
							System.out.println( cell.getStringCellValue() );
							break;
						case NUMERIC:
							System.out.println( cell.getNumericCellValue() );
							break;
						case BLANK:
							System.out.println( cell.getStringCellValue() + "(Blank)" );
							break;
						default:
							System.out.println( "#####" );
						}
					}
				}
			}
		}
		catch ( FileNotFoundException e ) {
		}
		catch ( IOException e ) {
		} finally {
			IOUtils.closeQuietly( workbook );
			IOUtils.closeQuietly( inputStream );
		}
	}
}
