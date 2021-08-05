package zcrb.excel.merging;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.ResourceBundle;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import zcrb.excel.merging.utils.UtfResourceBundle;

public class PoiMerging {

  private static List<String> tt = Arrays.asList(new String[] { "888999", "888999.0" });

  private static boolean isCellForAdd(Cell cell) {
    return cell != null && tt.contains(cell.toString());
  }

  private static class SrcDescr {
    File file;
    Workbook workBook;

    public SrcDescr(File newFile, Workbook newWorkbook) {
      this.file = newFile;
      this.workBook = newWorkbook;
    }

  }

  static UtfResourceBundle bundle = new UtfResourceBundle(ResourceBundle.getBundle("bundle"));

  public static void runAdder(Model model) throws Exception {

    model.sendMsg("==============");
    model.sendMsg("-=" + bundle.getString("PoiMerging.Msg.startMerge.text") + "=-");

    File templateFile = model.getTemplateFile();
    // Workbook templatewb = WorkbookFactory.create(new BufferedInputStream(new
    // FileInputStream(templateFile)));
    Workbook templatewb = WorkbookFactory.create(new BufferedInputStream(new FileInputStream(templateFile)));

    CellStyle debuggerCellStyle = templatewb.createCellStyle();
    debuggerCellStyle.setFillForegroundColor(IndexedColors.YELLOW.index);
    debuggerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    List<String> templateNameSheetList = new ArrayList<>();
    templatewb.sheetIterator().forEachRemaining(sheet -> templateNameSheetList.add(sheet.getSheetName()));

    File[] excelFiles = model.getExcelFiles();
    List<SrcDescr> srcDescrList = new ArrayList<>();
    // initialize list of srcDescr
    for (int q = 0; q < excelFiles.length; q++) {
      File currentFile = excelFiles[q];
      Workbook currentWorkbook = WorkbookFactory.create(new BufferedInputStream(new FileInputStream(currentFile)));
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
        throw new Exception("File:" + currentSrcDescr.file
            + bundle.getString("PoiMerging.Exception.notEqualSheets.text") + "(" + templateFile + ")");
      }
    }

    for (int q = 0; q < templateNameSheetList.size(); q++) {
      String currentSheetName = templateNameSheetList.get(q);
      Sheet curreSheetTemplatSheet = templatewb.getSheet(currentSheetName);

      for (int r = 0; r <= curreSheetTemplatSheet.getLastRowNum(); r++) {
        Row templateRow = curreSheetTemplatSheet.getRow(r);
        if (templateRow == null)
          continue;

        for (int c = 0; c < templateRow.getLastCellNum(); c++) {
          Cell cell = templateRow.getCell(c);
          boolean isCellForAdd = isCellForAdd(cell);
          if (isCellForAdd) {
            doAdd(model, srcDescrList, cell, debuggerCellStyle);
          }

        }

      }

    }

    SimpleDateFormat formatD = new SimpleDateFormat("dd.MM.yyyy HH-mm-ss");
    File srcDir = model.getSrcDir();
    File resultDir = new File(srcDir, "result");
    resultDir.mkdirs();
    File outFile = new File(resultDir, (model.isDebugMode() ? "DEBUGGER_VERSION_" : "") + formatD.format(new Date())
        + ".xls" + (templatewb instanceof XSSFWorkbook ? "x" : ""));
    OutputStream outStream = (new FileOutputStream(outFile));
    templatewb.write(outStream);
    outStream.close();

    model.sendMsg("-=" + bundle.getString("PoiMerging.Msg.FileSaved.text", outFile.toString()) + "=-");
  }

  private static void doAdd(Model model, List<SrcDescr> srcDescrList, Cell cell, CellStyle debugStyle) {
    boolean isDebugMode = model.isDebugMode();

    Workbook targetWb = cell.getSheet().getWorkbook();
    Sheet targetSheet = cell.getSheet();
    String sheetName = cell.getSheet().getSheetName();
    int rowIndex = cell.getRowIndex();
    int cellIndex = cell.getColumnIndex();

    double sum = 0;
    StringBuilder debuggerComment = new StringBuilder();

    for (int q = 0; q < srcDescrList.size(); q++) {
      SrcDescr srcDescr = srcDescrList.get(q);
      Workbook srcWB = srcDescr.workBook;

      Sheet srcSheet = srcWB.getSheet(sheetName);
      if (srcSheet == null) {
        model.sendMsg(bundle.getString("PoiMerging.Msg.SheetNotFound.text", "" + srcDescr.file, sheetName));
        continue;
      }
      Row srcRow = srcSheet.getRow(rowIndex);
      if (srcRow == null) {
        model.sendMsg(
            bundle.getString("PoiMerging.Msg.RowNotFound.text", "" + srcDescr.file, sheetName, "" + (rowIndex + 1)));
        continue;
      }
      Cell srcCell = srcRow.getCell(cellIndex);
      if (srcCell == null) {
        model.sendMsg(bundle.getString("PoiMerging.Msg.RowNotFound.text", "" + srcDescr.file, sheetName,
            "" + (rowIndex + 1), "" + cell.getAddress()));
        continue;
      }

      // Пустые строчки пропускаем и даже не сигналим о них
      if (srcCell.getCellType() == CellType.BLANK) {
        continue;
      }

      if (srcCell.getCellType() != CellType.NUMERIC) {

        model.sendMsg(bundle.getString("PoiMerging.Msg.CellHaveNotDigit.text", "" + srcDescr.file, sheetName,
            "" + (rowIndex + 1), "" + cellIndex, "" + cell.getAddress(), "" + srcCell));
        continue;
      }

      double srcValue = srcCell.getNumericCellValue();
      sum += srcValue;
      debuggerComment.append(srcDescr.file.getName() + " - " + srcValue).append('\n');
    }

    cell.setCellValue(sum);

    if (isDebugMode) {
      cell.setCellStyle(debugStyle);
      {
        cell.removeCellComment();
        CreationHelper factory = targetWb.getCreationHelper();
        ClientAnchor anchor = factory.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex()+ 20);
        anchor.setRow1(cell.getRowIndex());
        anchor.setRow2(cell.getRowIndex() + 10);


        Drawing<?> drawing = targetSheet.createDrawingPatriarch();
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(factory.createRichTextString(debuggerComment.toString()));
        cell.setCellComment(comment);

      }

    }
  }

}
