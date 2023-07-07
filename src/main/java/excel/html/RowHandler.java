package excel.html;

import org.apache.poi.ss.usermodel.Row;
import org.w3c.dom.Element;

/**
 * date: 2023/7/6 10:30
 * description:
 */
public class RowHandler implements IExcelHandler {

    private Row row;
    private DocumentHolder holder;

    public RowHandler(Row row, DocumentHolder holder) {
        this.row = row;
        this.holder = holder;
    }

    @Override
    public Element handle() {
        return holder.createTableRow();
    }
}
