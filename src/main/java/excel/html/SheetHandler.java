package excel.html;

import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Element;

/**
 * date: 2023/7/6 10:20
 * description:
 */
public class SheetHandler implements IExcelHandler {

    private Sheet sheet;
    private HtmlDocumentHolder holder;

    public SheetHandler(Sheet sheet, HtmlDocumentHolder holder) {
        this.sheet = sheet;
        this.holder = holder;
    }

    @Override
    public Element handle() {
        Element table = holder.createTable();
        return table;
    }
}
