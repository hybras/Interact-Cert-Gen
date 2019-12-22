package in.radiance.interact;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ExcelReader {

	public static void main(String[] args) {
		try {
			final XSSFWorkbook wb = new XSSFWorkbook("Attendance.xlsx");
			final XSSFSheet sh = wb.getSheetAt(0);
			final XSSFRow dateRow = sh.getRow(0);
			final XSSFRow eventRow = sh.getRow(1);
			final XSSFRow timeRow = sh.getRow(2);

			XWPFDocument template = new XWPFDocument(new FileInputStream("Template.docx"));

			for (int r = 3; r < sh.getLastRowNum() - 1; r++) {
				final XSSFRow record = sh.getRow(r);
				String first = record.getCell(0).toString();
				String last = record.getCell(1).toString();
				for (int c = 2; c < record.getLastCellNum(); c++) {
					String hours = record.getCell(c).toString();
					final String eventName = eventRow.getCell(c).toString();
					final String date = dateRow.getCell(c).toString();

					if (hours.equals("") || eventName.toLowerCase().contains("meeting"))
						continue;
					else if (hours.equals("t")) {
						hours = timeRow.getCell(c).toString();
					}

					// TODO Parse duration to double hours
					certify(template, String.join(" ", first, last), date.toString(), eventName, hours);
				}
			}
			template.close();
			wb.close();
		} catch (IOException e) {
		}
	}

	private static void certify(XWPFDocument template, String name, String date, String event, String hours) {
		replaceText(template, "DATE", date);
		replaceText(template, "NAME", name);
		replaceText(template, "EVENT", event);
		replaceText(template, "HOURS", hours + " hours");

		saveDocument(template, name + date + ".docx");
		// Undo Changes for next save
		replaceText(template, date, "DATE");
		replaceText(template, name, "NAME");
		replaceText(template, event, "EVENT");
		replaceText(template, hours + " hours", "HOURS");
	}

	public static void replaceText(XWPFDocument template, String find, String replace) {
		for (XWPFParagraph p : template.getParagraphs()) {
			List<XWPFRun> runs = p.getRuns();
			if (runs != null) {
				for (XWPFRun r : runs) {
					String text = r.getText(0);
					if (text != null && text.contains(find)) {
						text = text.replace(find, replace);// your content
						r.setText(text, 0);
					}
				}
			}
		}
	}

	private static void saveDocument(XWPFDocument template, String file) {
		try (FileOutputStream out = new FileOutputStream(file)) {
			template.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}