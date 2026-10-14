package fr.martine.suiviremodeling;
import android.app.Activity;
import android.os.Bundle;
import android.content.Intent;
import android.net.Uri;
import android.webkit.*;
import android.widget.Toast;
import java.nio.charset.StandardCharsets;
import java.io.OutputStream;
public class MainActivity extends Activity {
 WebView web; ValueCallback<Uri[]> picker; String pending;
 @Override public void onCreate(Bundle state) {
  super.onCreate(state); web=new WebView(this); setContentView(web);
  web.getSettings().setJavaScriptEnabled(true); web.getSettings().setDomStorageEnabled(true);
  web.getSettings().setAllowFileAccess(true); web.getSettings().setAllowContentAccess(true);
  web.addJavascriptInterface(new ExportBridge(),"AndroidExport");
  web.setWebViewClient(new WebViewClient(){@Override public boolean shouldOverrideUrlLoading(WebView v, WebResourceRequest r){return true;}});
  web.setWebChromeClient(new WebChromeClient(){
   @Override public boolean onShowFileChooser(WebView v,ValueCallback<Uri[]> cb,FileChooserParams params){
    if(picker!=null)picker.onReceiveValue(null); picker=cb;
    Intent i=new Intent(Intent.ACTION_OPEN_DOCUMENT);i.addCategory(Intent.CATEGORY_OPENABLE);i.setType("application/json"); startActivityForResult(i,2);return true;
   }
  });
  web.loadUrl("file:///android_asset/index.html");
 }
 public class ExportBridge {
  @JavascriptInterface public void save(String name,String text,String mime){
   runOnUiThread(()->{if(pending!=null){Toast.makeText(MainActivity.this,"Termine la sauvegarde en cours.",Toast.LENGTH_LONG).show();return;} pending=text;
    Intent i=new Intent(Intent.ACTION_CREATE_DOCUMENT);i.addCategory(Intent.CATEGORY_OPENABLE);i.setType(mime.split(";")[0]);i.putExtra(Intent.EXTRA_TITLE,name);
    startActivityForResult(i,1);
   });
  }
 }
 @Override protected void onActivityResult(int req,int result,Intent data){
  super.onActivityResult(req,result,data);
  if(req==2){if(picker!=null){picker.onReceiveValue(result==RESULT_OK&&data!=null?new Uri[]{data.getData()}:null);picker=null;}return;}
  if(req==1){String text=pending;pending=null;if(result==RESULT_OK&&data!=null&&text!=null){
   try(OutputStream out=getContentResolver().openOutputStream(data.getData())){out.write(text.getBytes(StandardCharsets.UTF_8));Toast.makeText(this,"Fichier enregistré.",Toast.LENGTH_LONG).show();}
   catch(Exception e){Toast.makeText(this,"Échec de la sauvegarde.",Toast.LENGTH_LONG).show();}
  }}
 }
 @Override public void onBackPressed(){web.evaluateJavascript("if(document.body.classList.contains('landing-mode')){true}else{goBack();false}",v->{if("true".equals(v))finish();});}
}