package bin.ellie.com.readexcel;

import java.util.ArrayList;
import java.util.List;

// singleton of all books

public class BooksList {

    private static BooksList instance = null;

    public static BooksList getInstance() {
        if(instance == null){
            instance = new BooksList();
        }
        return instance;
    }

    private List<MyBook> mBooks;

    private BooksList() {
        mBooks = new ArrayList<MyBook>();
    }

    public List<MyBook> getmBooks() {
        return mBooks;
    }
}
