package excel.html;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Element;

import java.util.Iterator;

/**
 * date: 2023/7/6 10:20
 * description:
 */
public class SheetHandler implements IExcelHandler {

    private Sheet sheet;
    private int firstColumn;
    private int endColumn;
    private boolean outputColumnHeader;

    private DocumentHolder holder;

    public SheetHandler(Sheet sheet, DocumentHolder holder, boolean outputColumnHeader) {
        this.sheet = sheet;
        this.holder = holder;
        this.outputColumnHeader = outputColumnHeader;
        getColumnBounds();
    }

    @Override
    public Element handle() {
        Element table = holder.createTable();
        table.setAttribute("class", "defaults");

        if (outputColumnHeader) {
            handleHeader(table);
        }
        return table;
    }

    private void handleHeader(Element table) {
        Element header = holder.createTableHeader();
        table.appendChild(header);
        Element row = holder.createTableRow();
        row.setAttribute("class", "rowHeader");
        header.appendChild(row);
        //noinspection UnusedDeclaration
        StringBuilder colName = new StringBuilder();
        for (int i = firstColumn; i < endColumn; i++) {
            colName.setLength(0);
            int cnum = i;
            do {
                colName.insert(0, (char) ('A' + cnum % 26));
                cnum /= 26;
            } while (cnum > 0);
            Element th = holder.createTableHeaderCell();
            th.setTextContent(colName.toString());
            th.setAttribute("class", "colHeader");
            row.appendChild(th);
        }
    }

    private void getColumnBounds() {
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
}
