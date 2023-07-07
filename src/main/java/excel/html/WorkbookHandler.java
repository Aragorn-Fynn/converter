package excel.html;

import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Element;

/**
 * date: 2023/7/6 10:04
 * description:
 */
public class WorkbookHandler implements IExcelHandler {

    private Workbook workbook;
    private DocumentHolder holder;

    public WorkbookHandler(Workbook workbook, DocumentHolder holder) {
        this.workbook = workbook;
        this.holder = holder;
    }

    @Override
    public Element handle() {
        return null;
    }
}
