package excel.html;

import org.apache.poi.ss.usermodel.Row;
import org.w3c.dom.Element;

/**
 * date: 2023/7/6 10:30
 * description:
 */
public class RowHandler implements IExcelHandler {

    private Row row;
    private boolean outputRowNum;
    private DocumentHolder holder;

    public RowHandler(Row row, DocumentHolder holder, boolean outputRowNum) {
        this.row = row;
        this.holder = holder;
        this.outputRowNum = outputRowNum;
    }

    @Override
    public Element handle() {
        Element tr = holder.createTableRow();
        if (outputRowNum) {
            Element rowNum = holder.createTableCell();
            rowNum.setTextContent(row.getRowNum()+1+"");
            rowNum.setAttribute("class", "rowHeader");
            tr.appendChild(rowNum);
        }

        return tr;

    }
}
