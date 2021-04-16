package zcrb.excel.adder;

import java.beans.PropertyChangeListener;
import java.beans.PropertyChangeSupport;
import java.io.File;
import java.io.FileFilter;

public class Model {

  private static final String TEMPLATE_FILE_REGEX = "^template\\.(xls|xlsx)?$";

  /**
   * Название событий
   */
  public static final String CAN_BE_RUN_CHANGED = "canBeRunChanged";
  public static final String MODEL_SEND_MESSAGE = "modelSendMessage";
  public static final String EXCEL_FILES_CHANGED = "excelFilesChanged";
  public static final String CHANGE_SRC_DIR = "changeScrDir";
  public static final String TEMPLATE_FILE_CHANGED = "templateFileChanged";

  private final PropertyChangeSupport pcs = new PropertyChangeSupport(this);

  public void addPropertyChangeListener(PropertyChangeListener listener) {
    this.pcs.addPropertyChangeListener(listener);
  }

  public void addPropertyChangeListener(String propertyName, PropertyChangeListener listener) {
    this.pcs.addPropertyChangeListener(propertyName, listener);
  }

  public void removePropertyChangeListener(PropertyChangeListener listener) {
    this.pcs.removePropertyChangeListener(listener);
  }

  public void sendMsg(String msg) {
    pcs.firePropertyChange(MODEL_SEND_MESSAGE, null, msg);
  }

  public Model() {

    this.pcs.addPropertyChangeListener(CHANGE_SRC_DIR, evt -> {

      FileFilter fileFilter = pathname -> {
        final String name = pathname.getName();
        if (name.matches(TEMPLATE_FILE_REGEX)) {
          return false;
        }
        return name.endsWith(".xls") || name.endsWith(".xlsx");
      };

      Model.this.setExcelFiles(Model.this.srcDir.listFiles(fileFilter));
    });

    this.pcs.addPropertyChangeListener(CHANGE_SRC_DIR, evt -> {
      File newSrcDir = (File) evt.getNewValue();
      File[] templateList = newSrcDir.listFiles((File dir, String name) -> {
        return name.matches(TEMPLATE_FILE_REGEX);
      });
      if (templateList.length == 0) {
        setTemplateFile(null);
      } else if (templateList.length > 1) {
        Model.this.sendMsg("В выбранной директории есть более одно файла Template: " + templateList.toString());
        setTemplateFile(null);
      } else {
        setTemplateFile(templateList[0]);
      }
    });

    this.pcs.addPropertyChangeListener(evt -> {
      if (evt.getPropertyName() == CAN_BE_RUN_CHANGED) {
        return;
      }
      setCanBeRun(templateFile != null && excelFiles != null && excelFiles.length > 0);
    });

  }

  // Директория которую нужно склеить
  private File srcDir = null;

  public File getSrcDir() {
    return srcDir;
  }

  public void setSrcDir(File newSrcDir) {
    if (!newSrcDir.isDirectory())
      throw new RuntimeException("Нужно выбрать директорию!");

    Object newValue = newSrcDir;
    this.srcDir = newSrcDir;
    pcs.firePropertyChange(CHANGE_SRC_DIR, null, newValue);
  }

  // список файлов из выбранной диерктории
  private File[] excelFiles = null;

  public File[] getExcelFiles() {
    return excelFiles;
  }

  public void setExcelFiles(File[] newExcelFiles) {
    this.excelFiles = newExcelFiles;
    pcs.firePropertyChange(EXCEL_FILES_CHANGED, null, excelFiles);
  }

  // Директория которую нужно склеить
  private File templateFile = null;

  public File getTemplateFile() {
    return templateFile;
  }

  public void setTemplateFile(File newTemplateFile) {
    this.templateFile = newTemplateFile;
    pcs.firePropertyChange(TEMPLATE_FILE_CHANGED, null, newTemplateFile);
  }

  private boolean isCanBeRun = false;

  public void setCanBeRun(boolean newIsCanBeRun) {
    this.isCanBeRun = newIsCanBeRun;
    pcs.firePropertyChange(CAN_BE_RUN_CHANGED, null, newIsCanBeRun);
  }

  public boolean isCanBeRun() {
    return isCanBeRun;
  }

  private boolean debugMode;

  public void setDebugMode(boolean newDubugMode) {
    this.debugMode = newDubugMode;
  }

  public boolean isDebugMode() {
    return this.debugMode;
  }

}
