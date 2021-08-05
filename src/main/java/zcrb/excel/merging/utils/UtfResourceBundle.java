package zcrb.excel.merging.utils;

import java.io.UnsupportedEncodingException;
import java.text.MessageFormat;
import java.util.ResourceBundle;

public class UtfResourceBundle {

  // feature variables
  private ResourceBundle bundle;
  private String fileEncoding;

  public UtfResourceBundle(ResourceBundle bundle) {
    this.bundle = bundle;
    this.fileEncoding = "UTF-8";
  }

  public String getString(String key, String... params) {
    return new MessageFormat(this.getString(key)).format(params);
  }

  public String getString(String key) {
    try {
      return new String(bundle.getString(key).getBytes("ISO-8859-1"), this.fileEncoding);
    } catch (UnsupportedEncodingException e) {
      e.printStackTrace();
      return key;
    }
  }
}