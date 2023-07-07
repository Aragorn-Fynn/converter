package excel.html;

import com.sun.org.apache.xml.internal.serializer.OutputPropertiesFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileWriter;
import java.io.StringWriter;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
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
    private DocumentHolder holder;

    /**
     * style handler
     */
    private StyleHandler styleHandler;

    /**
     * the first column number of cell that has value
     */
    private int firstColumn;

    /**
     * the last column number of cell that has value
     */
    private int endColumn;

    /**
     * merged ranges
     */
    private CellRangeAddress[][] mergedRanges;

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
        this.holder = new DocumentHolder();
        this.styleHandler = new StyleHandler();
        // 1. excel -> document
        convertWorkbook();

        // 2. doc -> string
        String res = serialize(holder.getDocument());

        return res;
    }

    private String serialize(Document document) throws Exception {
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        StreamResult out = new StreamResult(writer);

        transformer.setOutputProperty("encoding","utf-8");
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty( OutputKeys.METHOD, "html" );
        transformer.setOutputProperty(OutputPropertiesFactory.S_KEY_INDENT_AMOUNT, "2");
        transformer.setOutputProperty(OutputPropertiesFactory.S_KEY_LINE_SEPARATOR, "\n");
        transformer.transform(new DOMSource(document), out);
        return writer.getBuffer().toString();
    }

    private Element convertWorkbook() {
        Element workbookDoc = new WorkbookHandler(workbook, holder).handle();
        for (int idx=0; idx<workbook.getNumberOfSheets(); idx++) {
            Element sheetDoc = convertSheet(workbook.getSheetAt(idx));
            holder.getBody().appendChild(sheetDoc);
        }

        holder.getStyle().setTextContent(styleHandler.flushToString());
        return workbookDoc;
    }

    private Element convertSheet(Sheet sheet) {
        Element sheetDoc = new SheetHandler(sheet, holder, options.isOutputColumnHeader(), options.isOutputRowNum()).handle();
        Element body = holder.createTableBody();
        initColumnBoundsOfSheet(sheet);
        mergedRanges = buildMergedRangesMap(sheet);

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
        Element rowDoc = new RowHandler(row, holder, options.isOutputRowNum()).handle();
        String name = styleHandler.handle(workbook, row.getRowStyle());
        if (name != null) {
            rowDoc.setAttribute("class", name);
        }
        for (int columnNum=firstColumn; columnNum<endColumn; columnNum++) {
            CellRangeAddress range = getMergedRange(mergedRanges, row.getRowNum(), columnNum);
            if (range != null && (range.getFirstColumn() != columnNum || range.getFirstRow() != row.getRowNum())) {
                continue;
            }

            Cell cell = row.getCell(columnNum);
            Element cellDoc = convertCell(cell);
            if (range != null) {
                if (range.getFirstColumn() != range.getLastColumn()) {
                    cellDoc.setAttribute("colspan", (range.getLastColumn() - range.getFirstColumn() + 1+""));
                }
                if (range.getFirstRow() != range.getLastRow()) {
                    cellDoc.setAttribute("rowspan", (range.getLastRow() - range.getFirstRow() + 1+""));
                }
            }
            rowDoc.appendChild(cellDoc);
        }

        return rowDoc;
    }

    private Element convertCell(Cell cell) {
        Element cellDoc = new CellHandler(cell, holder).handle();
        if (cell != null) {
            String name = styleHandler.handle(workbook, cell.getCellStyle());
            if (name != null) {
                cellDoc.setAttribute("class", name);
            }
        }
        return cellDoc;
    }

    private void initColumnBoundsOfSheet(Sheet sheet) {
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

    private CellRangeAddress[][] buildMergedRangesMap(Sheet sheet) {
        CellRangeAddress[][] mergedRanges = new CellRangeAddress[1][];
        for (final CellRangeAddress cellRangeAddress : sheet.getMergedRegions()) {
            final int requiredHeight = cellRangeAddress.getLastRow() + 1;
            if (mergedRanges.length < requiredHeight) {
                mergedRanges = Arrays.copyOf(mergedRanges, requiredHeight, CellRangeAddress[][].class);
            }

            for (int r = cellRangeAddress.getFirstRow(); r <= cellRangeAddress.getLastRow(); r++) {
                final int requiredWidth = cellRangeAddress.getLastColumn() + 1;

                CellRangeAddress[] rowMerged = mergedRanges[r];
                if (rowMerged == null) {
                    rowMerged = new CellRangeAddress[requiredWidth];
                    mergedRanges[r] = rowMerged;
                } else {
                    final int rowMergedLength = rowMerged.length;
                    if (rowMergedLength < requiredWidth) {
                        rowMerged = mergedRanges[r] =
                                Arrays.copyOf(rowMerged, requiredWidth, CellRangeAddress[].class);
                    }
                }

                Arrays.fill(rowMerged, cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn() + 1, cellRangeAddress);
            }
        }
        return mergedRanges;
    }

    private CellRangeAddress getMergedRange(
            CellRangeAddress[][] mergedRanges, int rowNum, int columnNum) {
        CellRangeAddress[] mergedRangeRowInfo = rowNum < mergedRanges.length ? mergedRanges[rowNum]
                : null;

        return mergedRangeRowInfo != null
                && columnNum < mergedRangeRowInfo.length ? mergedRangeRowInfo[columnNum]
                : null;
    }

    public static void main(String[] args) throws Exception {
        Options options = new Options(false, true, true);
        ExcelToHtmlConverter converter = new ExcelToHtmlConverter(options);
        Files.list(Paths.get("D:/"))
                .filter(item -> item.getFileName().toString().endsWith("xlsx") || item.getFileName().toString().endsWith("xls"))
                .forEach(item -> {
                    try {
                        String html = converter.convert(item.toFile());
                        FileWriter writer = new FileWriter(new File(String.format("D:/test_%s.html", item.getFileName().toString())));
                        writer.append(html);
                        writer.close();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                });
    }
}
