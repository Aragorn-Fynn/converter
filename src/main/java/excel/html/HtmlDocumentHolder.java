package excel.html;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * date: 2023/7/5 13:24
 * description: tool for creating document.
 */
public class HtmlDocumentHolder {

    protected final Element body;
    protected final Document document;
    protected final Element head;
    protected final Element html;

    /**
     * Map from tag name, to map linking known styles and css class names
     */
    private Map<String, Map<String, String>> stylesheet = new LinkedHashMap<String, Map<String, String>>();

    public HtmlDocumentHolder(Document document) {
        this.document = document;

        html = document.createElement("html");
        document.appendChild(html);

        body = document.createElement("body");
        head = document.createElement("head");

        html.appendChild(head);
        html.appendChild(body);

        addStyleClass(body, "b", "white-space-collapsing:preserve;");
    }

    public void addStyleClass(Element element, String classNamePrefix,
                              String style) {
        String exising = element.getAttribute("class");
        String addition = getOrCreateCssClass(classNamePrefix, style);
        String newClassValue = isEmpty(exising) ? addition
                : (exising + " " + addition);
        element.setAttribute("class", newClassValue);
    }

    public Element createTable() {
        return document.createElement("table");
    }

    public Element createTableBody() {
        return document.createElement("tbody");
    }

    public Element createTableCell() {
        return document.createElement("td");
    }

    public Element createTableColumn() {
        return document.createElement("col");
    }

    public Element createTableColumnGroup() {
        return document.createElement("colgroup");
    }

    public Element createTableHeader() {
        return document.createElement("thead");
    }

    public Element createTableHeaderCell() {
        return document.createElement("th");
    }

    public Element createTableRow() {
        return document.createElement("tr");
    }

    public Text createText(String data) {
        return document.createTextNode(data);
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

    public String getOrCreateCssClass(String classNamePrefix, String style) {
        if (!stylesheet.containsKey(classNamePrefix)) {
            stylesheet.put(classNamePrefix, new LinkedHashMap<String, String>(
                    1));
        }

        Map<String, String> styleToClassName = stylesheet.get(classNamePrefix);
        String knownClass = styleToClassName.get(style);
        if (knownClass != null) {
            return knownClass;
        }

        String newClassName = classNamePrefix + (styleToClassName.size() + 1);
        styleToClassName.put(style, newClassName);
        return newClassName;
    }

    static boolean isEmpty(String str) {
        return str == null || str.length() == 0;
    }
}
