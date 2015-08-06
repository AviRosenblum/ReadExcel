package bin.ellie.com.readexcel;

import android.app.ProgressDialog;
import android.content.Context;
import android.os.AsyncTask;
import android.os.Environment;
import android.util.Log;

import org.apache.http.HttpResponse;
import org.apache.http.HttpStatus;
import org.apache.http.StatusLine;
import org.apache.http.client.ClientProtocolException;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.conn.params.ConnManagerParams;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.params.HttpConnectionParams;
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
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;


/** This class is using to parse Excel file from server into a local variable using jExcel jar.
    Parameters are Context for display dialog during parse,
    and URL of server for execute AsyncTask.

    Created by Avi Rozenblum
 **/


public class ReadExcel extends AsyncTask<String, Void, String> {

    private static final int REGISTRATION_TIMEOUT = 3 * 1000;
    private static final int WAIT_TIMEOUT = 30 * 1000;
    private final HttpClient httpclient = new DefaultHttpClient();
    final HttpParams params = httpclient.getParams();
    private HttpResponse response;
    private String content = null;
    private File xlFile;
    private ProgressDialog dialog;
    private Context context;

    public ReadExcel(Context context) {
        this.context = context;
    }

    @Override
    protected void onPreExecute() {
        super.onPreExecute();

        // display dialog while loading data
        dialog = new ProgressDialog(context);
        dialog.setMessage("Loading, Please wait");
        dialog.setCancelable(false);
        dialog.setCanceledOnTouchOutside(false);
        dialog.setProgressStyle(ProgressDialog.STYLE_SPINNER);
        dialog.show();
    }

    @Override
    protected String doInBackground(String... urls) {
        String URL = null;
        try {
            // get server url from params
            URL = urls[0];
            // establish connection to server
            HttpConnectionParams.setConnectionTimeout(params, REGISTRATION_TIMEOUT);
            HttpConnectionParams.setSoTimeout(params, WAIT_TIMEOUT);
            ConnManagerParams.setTimeout(params, WAIT_TIMEOUT);
            HttpGet httpGet = new HttpGet(URL);
            response = httpclient.execute(httpGet);
            StatusLine statusLine = response.getStatusLine();
            if (statusLine.getStatusCode() == HttpStatus.SC_OK) {
                // get content of Excel file to inputStream
                InputStream is = response.getEntity().getContent();
                // create new file
                xlFile = new File
                        (Environment.getExternalStorageDirectory(), "books.xls");
                FileOutputStream fos = new FileOutputStream(xlFile);
                int read = 0;
                byte[] buffer = new byte[1024];
                // write content from inputStream into the file
                while ((read = is.read(buffer)) > 0) {
                    fos.write(buffer, 0, read);
                }
                // close file outputStream
                if (fos != null) {
                    fos.close();
                    // close inputStream
                    if (is != null) {
                        is.close();
                    }
                }
                // clear singleton before loading data
                BooksList.getInstance().getmBooks().clear();
                // parse Excel from file into local singleton
                parseExcel(new FileInputStream(xlFile.getAbsolutePath()));
                // close connection to server
                response.getEntity().getContent().close();
            } else {
                Log.w("HTTP1:", statusLine.getReasonPhrase());
                response.getEntity().getContent().close();
                throw new IOException(statusLine.getReasonPhrase());
            }
        } catch (ClientProtocolException e) {
            Log.w("HTTP2:", e);
            content = e.getMessage();
            cancel(true);
        } catch (IOException e) {
            Log.w("HTTP3:", e);
            content = e.getMessage();
            cancel(true);
        } catch (Exception e) {
            Log.w("HTTP4:", e);
            content = e.getMessage();
            cancel(true);
        }
        return content;
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
    protected void onPostExecute(String s) {
        super.onPostExecute(s);
        // dismiss dialog
        dialog.dismiss();
    }

}
