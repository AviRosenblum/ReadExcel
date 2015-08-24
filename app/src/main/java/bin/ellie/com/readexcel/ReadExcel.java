package bin.ellie.com.readexcel;

import android.app.Activity;
import android.app.ProgressDialog;
import android.os.AsyncTask;
import android.os.Environment;
import android.util.Log;
import android.widget.Toast;

import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.params.HttpParams;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Iterator;

import javax.net.ssl.HttpsURLConnection;


/** This class is using to parse Excel file from server into a local variable using Apache POI.
    Parameters are Context for display dialog during parse,
    and URL of server for execute AsyncTask.

    Created by Avi Rozenblum
 **/


public class ReadExcel extends AsyncTask<String, Void, Boolean> {

    private static final int REGISTRATION_TIMEOUT = 3 * 1000;
    private static final int WAIT_TIMEOUT = 30 * 1000;
    private final HttpClient httpclient = new DefaultHttpClient();
    final HttpParams params = httpclient.getParams();
    private HttpResponse response;
    private Boolean result = false;
    private File xlFile;
    private ProgressDialog dialog;
    private Activity mActivity;

    // Constructor to getting the activity calling to this class to display dialog while downloading file.
    public ReadExcel(Activity activity) {
        this.mActivity = activity;
    }

    @Override
    protected void onPreExecute() {
        super.onPreExecute();

        // display dialog while loading data
        dialog = new ProgressDialog(mActivity);
        dialog.setMessage("Loading, Please wait");
        dialog.setCancelable(false);
        dialog.setCanceledOnTouchOutside(false);
        dialog.setProgressStyle(ProgressDialog.STYLE_SPINNER);
        dialog.show();
    }

    @Override
    protected Boolean doInBackground(String... urls) {
        URL url;
        try {
            url = new URL(urls[0]);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setReadTimeout(WAIT_TIMEOUT);
            conn.setConnectTimeout(REGISTRATION_TIMEOUT);
            conn.setRequestMethod("GET");
            conn.setDoInput(true);

            int responseCode = conn.getResponseCode();
            if (responseCode == HttpsURLConnection.HTTP_OK) {
                // get content of Excel file to inputStream
                InputStream is = conn.getInputStream();
                // create new file
                xlFile = new File
                        (Environment.getExternalStorageDirectory(), "books.xls");
                FileOutputStream fos = new FileOutputStream(xlFile);
                int read = 0;
                byte[] buffer = new byte[1024];
                // write content from inputStream to file
                while ((read = is.read(buffer)) > 0) {
                    fos.write(buffer, 0, read);
                }
                // close file outputStream
                fos.close();
                // close inputStream
                is.close();
                // clear singleton before loading data
                BooksList.getInstance().getmBooks().clear();
                // parse Excel from file into local singleton
                parseExcel(new FileInputStream(xlFile.getAbsolutePath()));
                // No Exceptions, result is true.
                result = true;
            } else {
                Log.d("Connection-Exception:", String.valueOf(conn.getResponseCode()) + conn.getResponseMessage());
                result = false;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }


    private void parseExcel(InputStream fis) {
        try {
            // set workBook from Excel file inputStream
            HSSFWorkbook myWorkBook = new HSSFWorkbook(fis);
            // get the first sheet
            HSSFSheet mySheet = myWorkBook.getSheetAt(0);
            Iterator<Row> rowIter = mySheet.rowIterator();
            while (rowIter.hasNext()) {
                // go over all rows and skip the first row
                HSSFRow myRow = (HSSFRow) rowIter.next();
                if (myRow.getRowNum() == 0) {
                    continue;
                }
                // create MyBook object to storage the book in the row
                MyBook mb = new MyBook("", "", "", "", "");
                Iterator<Cell> cellIter = myRow.cellIterator();
                // go over al columns in the row
                while (cellIter.hasNext()) {
                    HSSFCell myCell = (HSSFCell) cellIter.next();
                    String cellValue = "";
                    // check whether the type of the cell is numeric or string
                    if (myCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                        // get the content of the cell
                        cellValue = myCell.getStringCellValue();
                    } else {
                        if (myCell.getNumericCellValue() > 0) {
                            // get the content numeric of the cell
                            double d = myCell.getNumericCellValue();
                            // convert number to integer
                            int num = (int) d;
                            // set cell value with string of the number
                            cellValue = String.valueOf(num);
                        } else {
                            cellValue = "";
                        }
                    }
                    // set MyBook object by checking what is the
                    // column and set the content of cellValue to the appropriate parameter.
                    switch (myCell.getColumnIndex()) {
                        case 0:
                            mb.setName(cellValue);
                            break;
                        case 1:
                            mb.setContent(cellValue);
                            break;
                        case 2:
                            mb.setShall(cellValue);
                            break;
                        case 3:
                            mb.setArea(cellValue);
                            break;
                        case 4:
                            mb.setNote(cellValue);
                            break;
                        default:
                            break;
                    }
                }
                // add MyBook object to singleton
                BooksList.getInstance().getmBooks().add(mb);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    protected void onPostExecute(Boolean b) {
        super.onPostExecute(b);
        // dismiss dialog
        if (b) {
            Toast.makeText(mActivity, "All Done Successfully", Toast.LENGTH_SHORT).show();
        }
        dialog.dismiss();
    }

}
