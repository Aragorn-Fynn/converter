package excel.html;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatResult;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

/**
 * date: 2023/7/6 10:45
 * description:
 */
public class CellHandler implements IExcelHandler {

    private Cell cell;
    private HtmlDocumentHolder holder;

    public CellHandler(Cell cell, HtmlDocumentHolder holder) {
        this.cell = cell;
        this.holder = holder;
    }

    @Override
    public Element handle() {
        Element td = holder.createTableCell();

        if (cell != null) {
            CellStyle style = cell.getCellStyle();
            CellFormat cf = CellFormat.getInstance(
                    style.getDataFormatString());
            CellFormatResult result = cf.apply(cell);
            String content = result.text;
            if (content.equals("")) {
                content = "&nbsp;";
            }
            Text text = holder.createText(content);
            td.appendChild(text);
        }

        return td;
    }
}
