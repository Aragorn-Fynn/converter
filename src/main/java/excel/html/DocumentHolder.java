package excel.html;

import org.apache.poi.util.XMLHelper;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

/**
 * date: 2023/7/5 13:24
 * description: tool for creating document.
 */
public class DocumentHolder {

    /**
     * whole doc
     */
    protected final Document document;

    /**
     * html tag
     */
    protected final Element html;

    /**
     * head tag
     */
    protected final Element head;

    /**
     * style
     */
    private Element style;

    /**
     * body tag
     */
    protected final Element body;

    public DocumentHolder() throws Exception {
        this.document = XMLHelper.getDocumentBuilderFactory().newDocumentBuilder().newDocument();

        html = document.createElement("html");
        body = document.createElement("body");
        head = document.createElement("head");
        style = document.createElement( "style" );
        style.setAttribute( "type", "text/css" );

        document.appendChild(html);
        html.appendChild(head);
        html.appendChild(body);
        head.appendChild(style);
    }

    public Element getBody() {
        return body;
    }

    public Document getDocument() {
        return document;
    }

    public Element getHead() {
        return head;
    }

    public Element getStyle() {
        return style;
    }


    public Element createTable() {
        return document.createElement("table");
    }

    public Element createTableColumnGroup() {
        return document.createElement("colgroup");
    }

    public Element createTableColumn() {
        return document.createElement("col");
    }

    public Element createTableHeader() {
        return document.createElement("thead");
    }

    public Element createTableHeaderCell() {
        return document.createElement("th");
    }

    public Element createTableBody() {
        return document.createElement("tbody");
    }

    public Element createTableRow() {
        return document.createElement("tr");
    }

    public Element createTableCell() {
        return document.createElement("td");
    }

    public Text createText(String data) {
        return document.createTextNode(data);
    }

}
