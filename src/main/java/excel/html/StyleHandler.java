package excel.html;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * date: 2023/7/7 10:23
 * description: handle style
 */
public class StyleHandler {
    private static final Map<String, Map<String, String>> styleHolder = new LinkedHashMap<>();

    public StyleHandler() {
        Map<String, String> defaults = new HashMap<>();
        defaults.put("background-color", "white");
        defaults.put("color", "black");
        defaults.put("text-decoration", "none");
        defaults.put("direction", "ltr");
        defaults.put("text-transform", "none");
        defaults.put("text-ident", "0");
        defaults.put("letter-spacing","0");
        defaults.put("word-spacing","0");
        defaults.put("white-space","normal");
        defaults.put("unicode-bidi","normal");
        defaults.put("vertical-align","0");
        defaults.put("background-image","none");
        defaults.put("text-shadow","none");
        defaults.put("list-style-image","none");
        defaults.put("list-style-type","none");
        defaults.put("padding","0");
        defaults.put("margin","0");
        defaults.put("border-collapse","collapse");
        defaults.put("white-space","pre");
        defaults.put("vertical-align","bottom");
        defaults.put("font-style","normal");
        defaults.put("font-family","sans-serif");
        defaults.put("font-variant","normal");
        defaults.put("font-weight","normal");
        defaults.put("font-size","10pt");
        defaults.put("text-align","right");
        styleHolder.put("defaults", defaults);

        Map<String, String> td = new HashMap<>();
        td.put("padding","1px 5px");
        td.put("border","1px solid silver");
        styleHolder.put("td", td);

        Map<String, String> colHeader = new HashMap<>();
        colHeader.put("background-color","silver");
        colHeader.put("font-weight","bold");
        colHeader.put("border","1px solid black");
        colHeader.put("text-align","center");
        colHeader.put("padding","1px 5px");
        styleHolder.put("colHeader", colHeader);

        Map<String, String> rowHeader = new HashMap<>();
        rowHeader.put("background-color","silver");
        rowHeader.put("font-weight","bold");
        rowHeader.put("border","1px solid black");
        rowHeader.put("text-align","right");
        rowHeader.put("padding","1px 5px");
        styleHolder.put("rowHeader", rowHeader);
    }

    public String handle(Workbook workbook, CellStyle style) {
        if (style == null) {
            return null;
        }

        String name = generateName(style);
        if (styleHolder.containsKey(name)) {
            return name;
        }

        Map<String, String> styleMap = new HashMap<>();
        styleHolder.put(name, styleMap);
        handleAlign(style, styleMap);
        handleFont(workbook, style, styleMap);
        handleBorder(style, styleMap);
        handleColor(workbook, style, styleMap);

        return name;
    }

    private void handleColor(Workbook workbook, CellStyle style, Map<String, String> styleMap) {
        if (style instanceof XSSFCellStyle) {
            XSSFCellStyle cs = (XSSFCellStyle) style;
            styleMap.put("background-color", styleColor(cs.getFillForegroundXSSFColor()));
            styleMap.put("text-color", styleColor(cs.getFont().getXSSFColor()));
        } else if (style instanceof HSSFCellStyle) {
            HSSFCellStyle cs = (HSSFCellStyle) style;
            HSSFPalette colors = ((HSSFWorkbook) workbook).getCustomPalette();
            styleMap.put("background-color", styleColor(colors.getColor(cs.getFillForegroundColor())));
            styleMap.put("color", styleColor(colors.getColor(cs.getFont(workbook).getColor())));
            styleMap.put("border-left-color", styleColor(colors.getColor(cs.getLeftBorderColor())));
            styleMap.put("border-right-color", styleColor(colors.getColor(cs.getRightBorderColor())));
            styleMap.put("border-top-color", styleColor(colors.getColor(cs.getTopBorderColor())));
            styleMap.put("border-bottom-color", styleColor(colors.getColor(cs.getBottomBorderColor())));
        }
    }

    private String styleColor(HSSFColor color) {
        short[] rgb = color.getTriplet();
        if (color == null || color.getIndex() == HSSFColor.HSSFColorPredefined.AUTOMATIC.getColor().getIndex()) {
            return "";
        }
        return String.format("#%02x%02x%02x", rgb[0], rgb[1], rgb[2]);
    }

    private String styleColor(XSSFColor color) {
        String res = "";
        if (color == null || color.isAuto()) {
            return res;
        }

        byte[] rgb = color.getRGB();
        if (rgb == null) {
            return res;
        }

        return String.format("#%02x%02x%02x", rgb[0], rgb[1], rgb[2]);
    }

    private void handleBorder(CellStyle style, Map<String, String> styleMap) {
        BorderStyle left = style.getBorderLeft();
        if (BORDER.containsKey(left)) {
            styleMap.put("border-left", BORDER.get(left));
        }

        BorderStyle right = style.getBorderRight();
        if (BORDER.containsKey(right)) {
            styleMap.put("border-right", BORDER.get(right));
        }

        BorderStyle top = style.getBorderTop();
        if (BORDER.containsKey(top)) {
            styleMap.put("border-top", BORDER.get(top));
        }

        BorderStyle bottom = style.getBorderBottom();
        if (BORDER.containsKey(bottom)) {
            styleMap.put("border-bottom", BORDER.get(bottom));
        }
    }

    private void handleFont(Workbook workbook, CellStyle style, Map<String, String> styleMap) {
        Font font = workbook.getFontAt(style.getFontIndexAsInt());

        if (font.getBold()) {
            styleMap.put("font-weight", "bold");
        }
        if (font.getItalic()) {
            styleMap.put("font-style", "italic");
        }

        int fontheight = font.getFontHeightInPoints();
        styleMap.put("font-size", fontheight+"pt");
    }

    private void handleAlign(CellStyle style, Map<String, String> styleMap) {
        HorizontalAlignment alignment = style.getAlignment();
        if (HALIGN.containsKey(alignment)) {
            styleMap.put("text-align", HALIGN.get(alignment));
        }

        VerticalAlignment verticalAlignment = style.getVerticalAlignment();
        if (VALIGN.containsKey(verticalAlignment)) {
            styleMap.put("vertical-align", VALIGN.get(verticalAlignment));
        }
    }

    private String generateName(CellStyle style) {
        return "style_" + style.getIndex();
    }

    public String flushToString() {
        StringBuffer sb = new StringBuffer();
        for (String name : styleHolder.keySet()) {
            sb.append("\n").append("defaults".equals(name) ? ".defaults":(".defaults ."+name)).append(" {\n");
            styleHolder.get(name).entrySet().forEach(
                    item -> {
                        sb.append(item.getKey()).append(" : ").append(item.getValue()).append(";\n");
                    }
            );
            sb.append("}\n");
        }

        return sb.toString();
    }

    private static final Map<HorizontalAlignment, String> HALIGN = mapFor(
            HorizontalAlignment.LEFT, "left",
            HorizontalAlignment.CENTER, "center",
            HorizontalAlignment.RIGHT, "right",
            HorizontalAlignment.FILL, "left",
            HorizontalAlignment.JUSTIFY, "left",
            HorizontalAlignment.CENTER_SELECTION, "center");

    private static final Map<VerticalAlignment, String> VALIGN = mapFor(
            VerticalAlignment.BOTTOM, "bottom",
            VerticalAlignment.CENTER, "middle",
            VerticalAlignment.TOP, "top");

    private static final Map<BorderStyle, String> BORDER = mapFor(
            BorderStyle.DASH_DOT, "dashed 1pt",
            BorderStyle.DASH_DOT_DOT, "dashed 1pt",
            BorderStyle.DASHED, "dashed 1pt",
            BorderStyle.DOTTED, "dotted 1pt",
            BorderStyle.DOUBLE, "double 3pt",
            BorderStyle.HAIR, "solid 1px",
            BorderStyle.MEDIUM, "solid 2pt",
            BorderStyle.MEDIUM_DASH_DOT, "dashed 2pt",
            BorderStyle.MEDIUM_DASH_DOT_DOT, "dashed 2pt",
            BorderStyle.MEDIUM_DASHED, "dashed 2pt",
            BorderStyle.NONE, "none",
            BorderStyle.SLANTED_DASH_DOT, "dashed 2pt",
            BorderStyle.THICK, "solid 3pt",
            BorderStyle.THIN, "dashed 1pt");

    @SuppressWarnings({"unchecked"})
    private static <K, V> Map<K, V> mapFor(Object... mapping) {
        Map<K, V> map = new HashMap<K, V>();
        for (int i = 0; i < mapping.length; i += 2) {
            map.put((K) mapping[i], (V) mapping[i + 1]);
        }
        return map;
    }
}
