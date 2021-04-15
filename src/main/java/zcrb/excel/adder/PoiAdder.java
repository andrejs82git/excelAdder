package zcrb.excel.adder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PoiAdder {

  private static List<String> tt = Arrays.asList(new String[] { "888999", "888999.0" });

  private static boolean isCellForAdd(Cell cell) {
    return tt.contains(cell.toString());
  }

  private static class SrcDescr {
    File file;
    Workbook workBook;

    public SrcDescr(File newFile, Workbook newWorkbook) {
      this.file = newFile;
      this.workBook = newWorkbook;
    }

  }

  public static void runAdder(Model model) throws Exception {
    model.sendMsg("==============");
    model.sendMsg("-=Старт склейки файлов в один.=-");


    File templateFile = model.getTemplateFile();
    Workbook templatewb = WorkbookFactory.create(new FileInputStream(templateFile));
    // Workbook targetwb = WorkbookFactory.create(new
    // FileInputStream(templateFile));

    List<String> templateNameSheetList = new ArrayList<>();
    templatewb.sheetIterator().forEachRemaining(sheet -> templateNameSheetList.add(sheet.getSheetName()));

    File[] excelFiles = model.getExcelFiles();
    List<SrcDescr> srcDescrList = new ArrayList<>();
    // initialize list of srcDescr
    for (int q = 0; q < excelFiles.length; q++) {
      File currentFile = excelFiles[q];
      Workbook currentWorkbook = WorkbookFactory.create(new FileInputStream(currentFile));
      SrcDescr srcDescr = new SrcDescr(currentFile, currentWorkbook);
      srcDescrList.add(srcDescr);
    }

    // test for all sheets
    for (int q = 0; q < srcDescrList.size(); q++) {
      SrcDescr currentSrcDescr = srcDescrList.get(q);
      Workbook currentWorkbook = currentSrcDescr.workBook;
      List<String> sheetNameList = new ArrayList<>();
      currentWorkbook.sheetIterator().forEachRemaining(sheet -> sheetNameList.add(sheet.getSheetName()));
      if (!templateNameSheetList.containsAll(sheetNameList)) {
        throw new Exception("Файл:" + currentSrcDescr.file
            + " не содержит все листы которые присутсвуют в шаблонном файле (" + templateFile + ")");
      }
    }

    for (int q = 0; q < templateNameSheetList.size(); q++) {
      String currentSheetName = templateNameSheetList.get(q);
      Sheet curreSheetTemplatSheet = templatewb.getSheet(currentSheetName);

      for (int r = 0; r <= curreSheetTemplatSheet.getLastRowNum(); r++) {
        Row templateRow = curreSheetTemplatSheet.getRow(r);

        for (int c = 0; c < templateRow.getLastCellNum(); c++) {
          Cell cell = templateRow.getCell(c);
          boolean isCellForAdd = isCellForAdd(cell);
          if (isCellForAdd) {
            doAdd(model, srcDescrList, cell);
          }

        }

      }

    }

    SimpleDateFormat formatD = new SimpleDateFormat("dd.MM.yyyy HH-mm-ss");
    File srcDir = model.getSrcDir();
    File resultDir = new File(srcDir, "result");
    resultDir.mkdirs();
    File outFile = new File(resultDir, formatD.format(new Date()) + ".xls");
    OutputStream outStream = (new FileOutputStream(outFile));
    templatewb.write(outStream);
    outStream.close();

    model.sendMsg("-=Файл (" + outFile + ") сохранен.=-");
  }

  private static void doAdd(Model model, List<SrcDescr> srcDescrList, Cell cell) {
    boolean isDebugMode = model.isDebugMode();

    String sheetName = cell.getSheet().getSheetName();
    int rowIndex = cell.getRowIndex();
    int cellIndex = cell.getColumnIndex();


    double sum = 0;
    for (int q = 0; q < srcDescrList.size(); q++) {
      SrcDescr srcDescr = srcDescrList.get(q);
      Workbook srcWB = srcDescr.workBook;

      Sheet srcSheet = srcWB.getSheet(sheetName);
      Row srcRow = srcSheet.getRow(rowIndex);
      if (srcRow == null) {
        model.sendMsg("Файл (" + srcDescr.file + ") Лист (" + sheetName + ") строчка (" + (rowIndex + 1)
            + ") не найдена! (ПРОПУСКАЮ)");
        continue;
      }
      Cell srcCell = srcRow.getCell(cellIndex);
      if (srcCell == null) {
        model.sendMsg("Файл (" + srcDescr.file + ") Лист (" + sheetName + ") строчка (" + (rowIndex + 1) + ") ячейка("
            + cellIndex + ") не найдена! (ПРОПУСКАЮ)");
        continue;
      }
      double srcValue = srcCell.getNumericCellValue();
      sum += srcValue;
    }

    cell.setCellValue(sum);

    if (isDebugMode) {
      // HSSFPatriarch hpt = (HSSFPatriarch) cell.getSheet().createDrawingPatriarch();
      // cell.setCellComment();
    }
  }

}
