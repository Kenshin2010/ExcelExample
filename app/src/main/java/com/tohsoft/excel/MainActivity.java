package com.tohsoft.excel;

import android.Manifest;
import android.annotation.SuppressLint;
import android.net.Uri;
import android.os.AsyncTask;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.widget.TabHost;
import android.widget.TextView;

import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;

import com.tbruyelle.rxpermissions2.Permission;
import com.tbruyelle.rxpermissions2.RxPermissions;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import io.reactivex.functions.Consumer;
import jxl.Sheet;
import jxl.Workbook;

public class MainActivity extends AppCompatActivity {

    File file;
    Uri uri;
    TabHost tabHost;

    @RequiresApi(api = Build.VERSION_CODES.JELLY_BEAN) @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        new RxPermissions(this).requestEach(Manifest.permission.READ_EXTERNAL_STORAGE)
                .subscribe(new Consumer<Permission>() {
                    @Override
                    public void accept(Permission permission) throws Exception {
                        if (permission.granted) {

                        }
                    }
                });
        tabHost = findViewById(R.id.sheets);
        new XlsAsyncTask().execute();
    }

    @SuppressLint("StaticFieldLeak")
    public class XlsAsyncTask extends AsyncTask<String, String, String> {
        private XlsAsyncTask() {
        }

        /* access modifiers changed from: protected */
        public void onPreExecute() {
            super.onPreExecute();
        }

        /* access modifiers changed from: protected */
        public String doInBackground(String... strArr) {
            if (Constutil.f8593f != null) {
                file = Constutil.f8593f;
            }
//            else if (uri != null) {
//                uri = getIntent().getData();
//                if (uri != null) {
//                    if (uri == null || !"content".equals(uri.getScheme())) {
//                        Constutil.f8593f = new File(uri.getPath());
//                        file = Constutil.f8593f;
//                    } else {
//                        Cursor query = getContentResolver().query(uri, new String[]{"_data"}, null, null, null);
//                        query.moveToFirst();
//                        Constutil.f8593f = new File(query.getString(0));
//                        query.close();
//                        file = Constutil.f8593f;
            file = new File(Environment.getExternalStorageDirectory() + File.separator + "Download" + File.separator + "PDF Reader Functions.xlsx");
//                    }
//                }
//            }
            if (file.getName().endsWith(Constant.excelExtension)) {
                readXlsFile(file);
                return "xls";
            } else {
                readXlsxFile(file);
                return "xls";
            }
        }

        /* access modifiers changed from: protected */
        public void onPostExecute(String str) {
            super.onPostExecute(str);
            if (tabHost != null && tabHost.getTabWidget() != null) {
                for (int i = 0; i < tabHost.getTabWidget().getChildCount(); i++) {
                    tabHost.getTabWidget().getChildAt(i);
                    ((TextView) tabHost.getTabWidget().getChildAt(i).findViewById(android.R.id.title)).setTextColor(getResources().getColor(R.color.black));
                }
            }
        }
    }

    public void readXlsFile(final File file2) {
        runOnUiThread(new Runnable() {
            @Override public void run() {
                try {
                    Workbook workbook = Workbook.getWorkbook(new FileInputStream(file2));
                    tabHost.setup();
                    for (final Sheet sheet : workbook.getSheets()) {
                        TabHost.TabSpec newTabSpec = tabHost.newTabSpec(sheet.getName());
                        newTabSpec.setContent(new TabHost.TabContentFactory() {
                            public View createTabContent(String str) {
                                final XlsSheetView xlsSheetView = new XlsSheetView(MainActivity.this);
                                xlsSheetView.setSheet(sheet);
                                return xlsSheetView;
                            }
                        });
                        newTabSpec.setIndicator("    " + sheet.getName() + "    ");
                        tabHost.addTab(newTabSpec);
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }

            }
        });
    }

    /* access modifiers changed from: private */
    public void readXlsxFile(File file2) {
        try {
            XSSFWorkbook xSSFWorkbook = new XSSFWorkbook(new FileInputStream(file2));
            HSSFWorkbook hSSFWorkbook = new HSSFWorkbook();
            int numberOfSheets = xSSFWorkbook.getNumberOfSheets();
            for (int i = 0; i < numberOfSheets; i++) {
                XSSFSheet sheetAt = xSSFWorkbook.getSheetAt(i);
                org.apache.poi.ss.usermodel.Sheet createSheet = hSSFWorkbook.createSheet(sheetAt.getSheetName());
                Iterator<Row> rowIterator = sheetAt.rowIterator();
                while (rowIterator.hasNext()) {
                    Row next = rowIterator.next();
                    Row createRow = createSheet.createRow(next.getRowNum());
                    Iterator<Cell> cellIterator = next.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell next2 = cellIterator.next();
                        Cell createCell = createRow.createCell(next2.getColumnIndex(), next2.getCellType());
                        switch (next2.getCellType()) {
                            case 0:
                                createCell.setCellValue(next2.getNumericCellValue());
                                break;
                            case 1:
                                createCell.setCellValue(next2.getStringCellValue());
                                break;
                            case 2:
                                createCell.setCellFormula(next2.getCellFormula());
                                break;
                            case 4:
                                createCell.setCellValue(next2.getBooleanCellValue());
                                break;
                            case 5:
                                createCell.setCellValue((double) next2.getErrorCellValue());
                                break;
                        }
                        createCell.getCellStyle().setDataFormat(next2.getCellStyle().getDataFormat());
                        createCell.setCellComment(next2.getCellComment());
                    }
                }
            }
            File createTempFile = File.createTempFile("myTempXlsFile", Constant.excelExtension, getApplicationContext().getCacheDir());
            BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(new FileOutputStream(createTempFile));
            hSSFWorkbook.write(bufferedOutputStream);
            bufferedOutputStream.close();
            readXlsFile(createTempFile);
        } catch (Exception e) {
            e.printStackTrace();
            Log.e("hi", "" + e);
        }
    }


}
