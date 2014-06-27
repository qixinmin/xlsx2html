package temp.xlsx;

import android.app.Activity;
import android.app.ProgressDialog;
import android.os.AsyncTask;
import android.os.Bundle;
import android.os.Handler;
import android.text.TextUtils;
import android.util.Log;
import android.webkit.WebSettings;
import android.webkit.WebView;

public class XlsxShow extends Activity {
    final String TAG = "XlsxShow";
    Handler handler = new Handler();
    ProgressDialog pd;
    WebView webView;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.xlsx_show);
        webView = (WebView) findViewById(R.id.webview);
        final String temp = getIntent().getStringExtra("dst");
        final String path;
        if (TextUtils.isEmpty(temp)) {
            path = "/sdcard/aaa/xlsx/Chart.xlsx";
        } else {
            path = temp;
        }
        Log.d(TAG,"onCreate"+path);
        setTitle(path);
        AsyncTask<String,Integer,String> at = new AsyncTask<String, Integer, String>() {
            @Override
            protected void onPreExecute(){
//                pd=new ProgressDialog(XlsxShow.this);
//                pd.setMessage("正在对文件进行解析...");
//                pd.setCanceledOnTouchOutside(false);
//                pd.show();
            }

            @Override
            protected String doInBackground(String... p) {
                Log.d(TAG,"doInbackground() ");
                String html = "<p>null<p>";
                try{
                    Xlsx2Html x2h = new Xlsx2Html(path);
                    html = x2h.convert();
//                    html = ToHtml.main(new String[]{path,path+".html"});//x2h.convert();
//                    html = "<html><body><h1>"+" convert success!"+"</h1></body></html>";
                }catch (Exception e){
                    html = "<html><body><h5> (XlsxShow)<br>"+e.toString()+"</h5></body></html>";
                }
                Log.d(TAG,html);
                return html;
            }
            @Override
            protected void onPostExecute(String html){
//                pd.dismiss();
                if (webView != null) {
                    WebSettings webSettings = webView.getSettings();
                    webSettings.setJavaScriptEnabled(true);
                    webSettings.setSupportZoom(true);
                    webSettings.setBuiltInZoomControls(true);
                    webSettings.setDisplayZoomControls(false);//隐藏缩放按钮
                    webSettings.setUseWideViewPort(true);
                    webSettings.setLoadWithOverviewMode(true);
                    webView.loadDataWithBaseURL("file://"+XlsxShow.this.getFilesDir().getAbsolutePath()+"/", html, "text/html", "UTF-8", null);
                } else {
                    Log.e(TAG, "web view is null!");
                }
            }
        };

        at.execute(path);

    }


}
