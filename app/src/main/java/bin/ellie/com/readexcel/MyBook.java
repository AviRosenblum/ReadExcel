package bin.ellie.com.readexcel;


public class MyBook {

    private String name;
    private String shall;
    private String area;
    private String note;
    private String content;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getShall() {
        return shall;
    }

    public void setShall(String shall) {
        this.shall = shall;
    }

    public String getArea() {
        return area;
    }

    public void setArea(String area) {
        this.area = area;
    }

    public String getNote() {
        return note;
    }

    public void setNote(String note) {
        this.note = note;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public MyBook(String name, String shall, String area, String note, String content) {
        this.name = name;
        this.shall = shall;
        this.area = area;
        this.note = note;
        this.content = content;
    }
}
