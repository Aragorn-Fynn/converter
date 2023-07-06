package excel.html;

import com.sun.org.apache.xml.internal.serialize.OutputFormat;
import com.sun.org.apache.xml.internal.serialize.XMLSerializer;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.XMLHelper;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.File;
import java.io.FileWriter;
import java.io.StringWriter;
import java.util.Iterator;

/**
 * date: 2023/7/5 10:50
 * description: A tool for converting excel to html using poi.
 */
public class ExcelToHtmlConverter {

    /**
     * input
     */
    private Workbook workbook;

    /**
     * output
     */
    private StringWriter writer;

    /**
     * options of output
     */
    private Options options;

    /**
     * html doc holder
     */
    private HtmlDocumentHolder holder;

    /**
     * the first column number of cell that has value
     */
    private int firstColumn;

    /**
     * the last column number of cell that has value
     */
    private int endColumn;
    /**
     * Creates a new converter
     */
    public ExcelToHtmlConverter(Options options) {
        this.options = options;
    }

    public String convert(String path) throws Exception {
        return convert(new File(path));
    }

    public String convert(File file) throws Exception {
        return convert(WorkbookFactory.create(file));
    }

    public String convert(Workbook workbook) throws Exception {
        this.workbook = workbook;
        this.writer = new StringWriter();
        this.holder = new HtmlDocumentHolder(XMLHelper.getDocumentBuilderFactory().newDocumentBuilder().newDocument());

        // 1. excel -> document
        convertWorkbook();

        // 2. doc -> string
        String res = serialize(holder.getDocument());

        return res;
    }

    private String serialize(Document document) throws Exception {
        OutputFormat format = new OutputFormat(document);
        XMLSerializer serializer = new XMLSerializer(writer, format);
        serializer.asDOMSerializer();
        serializer.serialize(document);
        return writer.getBuffer().toString();
    }

    private Element convertWorkbook() {
        Element workbookDoc = new WorkbookHandler(workbook, holder).handle();
        for (int idx=0; idx<workbook.getNumberOfSheets(); idx++) {
            Element sheetDoc = convertSheet(workbook.getSheetAt(idx));
            holder.getBody().appendChild(sheetDoc);
        }
        return workbookDoc;
    }

    private Element convertSheet(Sheet sheet) {
        Element sheetDoc = new SheetHandler(sheet, holder).handle();

        Element body = holder.createTableBody();
        getColumnBounds(sheet);
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Element rowDoc = convertRow(row);
            body.appendChild(rowDoc);
        }

        sheetDoc.appendChild(body);
        return sheetDoc;
    }

    private Element convertRow(Row row) {
        Element rowDoc = new RowHandler(row, holder).handle();
        for (int columnNum=firstColumn; columnNum<endColumn; columnNum++) {
            Cell cell = row.getCell(columnNum);
            Element cellDoc = convertCell(cell);
            rowDoc.appendChild(cellDoc);
        }

        return rowDoc;
    }

    private Element convertCell(Cell cell) {
        Element cellDoc = new CellHandler(cell, holder).handle();
        return cellDoc;
    }

    private void getColumnBounds(Sheet sheet) {
        Iterator<Row> iter = sheet.rowIterator();
        firstColumn = (iter.hasNext() ? Integer.MAX_VALUE : 0);
        endColumn = 0;
        while (iter.hasNext()) {
            Row row = iter.next();
            short firstCell = row.getFirstCellNum();
            if (firstCell >= 0) {
                firstColumn = Math.min(firstColumn, firstCell);
                endColumn = Math.max(endColumn, row.getLastCellNum());
            }
        }
    }

    public static void main(String[] args) throws Exception {
        String html = new ExcelToHtmlConverter(new Options()).convert("D:/1043826_11_3faf0317eac73bc4a89f18a802e369d.xlsx");
        FileWriter writer = new FileWriter(new File("D:/test.html"));
        writer.append(html);
        writer.close();
    }
}
