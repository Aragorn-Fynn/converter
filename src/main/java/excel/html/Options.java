package excel.html;

/**
 * date: 2023/7/5 11:19
 * description: options for converter
 */
public class Options {
    /**
     * 是否生成完整的html
     */
    private boolean outputCompleteHtml;
    /**
     * 是否输出列头：A,B,C
     */
    private boolean outputColumnHeader;
    /**
     * 是否输出行号
     */
    private boolean outputRowNum;

    public Options(boolean outputCompleteHtml, boolean outputColumnHeader, boolean outputRowNum) {
        this.outputCompleteHtml = outputCompleteHtml;
        this.outputColumnHeader = outputColumnHeader;
        this.outputRowNum = outputRowNum;
    }

    public boolean isOutputCompleteHtml() {
        return outputCompleteHtml;
    }

    public boolean isOutputColumnHeader() {
        return outputColumnHeader;
    }

    public boolean isOutputRowNum() {
        return outputRowNum;
    }
}
