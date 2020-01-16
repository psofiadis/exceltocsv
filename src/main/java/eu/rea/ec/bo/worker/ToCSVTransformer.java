package eu.rea.ec.bo.worker;

import static eu.rea.ec.bo.util.Utils.SEPARATOR_REPLACEMENT;
import static org.apache.poi.ss.usermodel.CellType.FORMULA;

import eu.rea.ec.bo.exception.ExcelToCSVGenericException;
import eu.rea.ec.bo.util.Utils;
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.tools.ant.DirectoryScanner;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class ToCSVTransformer {
  private static final Logger LOG = LogManager.getLogger(ToCSVTransformer.class);

  private String abacRootPathDir;
  private String abacRelativeTransformDir;
  private String abacRelativeProcessedDir;
  private String _abacProcessedDir="processed";
  private String abacFullPathCopyDirs;
  private String[] _abacFullPathCopyDirs;
  private String excelFilePatterns;
  private String[] _excelFilePatterns = new String[]{"**.xls", "**.xlsx"};

  private String _generateFilePerTab;
  private boolean generateFilePerTab = false;

  private String csvSeparator;
  private String _csvSeparator = ",";
  private String excelSheetRange;
  private int[] _excelSheetRange =new int[]{0,100};
  private String rowStart;
  private int _rowStart = 0;
  private String columnRange;
  private int[] _columnRange = new int[]{0,100};
  private DataFormatter formatter = new DataFormatter(true);
  private FormulaEvaluator evaluator;

  public void execute()  {
    validate();

    List<File> filenames = getExcelFilesForTransform();
    for (File file : filenames) {
      try {
        String rootDestinationFileName = file.getName();
        if(!this.generateFilePerTab){
          String destinationFilename = rootDestinationFileName;
          Workbook workbook = openWorkbook(file);
          evaluator = workbook.getCreationHelper().createFormulaEvaluator();
          List<List<String>> csvData = this.convertToCSV(workbook);
//        String date = Utils.getGeneratedDate();
          String csvFilename = destinationFilename.substring(0, destinationFilename.lastIndexOf('.')) + ".csv";
          File csvFile = new File(abacRootPathDir + File.separator + abacRelativeTransformDir, csvFilename);
          csvFile.createNewFile();
          LOG.info("Saving file " + csvFile.getName());
          this.saveCSVFile(csvFile, csvData);
          LOG.info("Saved file " + csvFile.getName());

          file.renameTo(new File(abacRootPathDir + File.separator + _abacProcessedDir + File.separator + file.getName()));
        }else{
          String destinationFilename = rootDestinationFileName;
          Workbook workbook = openWorkbook(file);
          evaluator = workbook.getCreationHelper().createFormulaEvaluator();
          List<Sheet> worksheets =  this.getWorksheets(workbook);
          for (Sheet worksheet : worksheets) {
            List<List<String>> csvData = this.convertToCSV(worksheet);
            String csvFilename = destinationFilename.substring(0, destinationFilename.lastIndexOf('.')) + "_" + worksheet
                .getSheetName() +
                ".csv";
            File csvFile = new File(abacRootPathDir + File.separator + abacRelativeTransformDir, csvFilename);
            csvFile.createNewFile();
            LOG.info("Saving file " + csvFile.getName());
            this.saveCSVFile(csvFile, csvData);
            LOG.info("Saved file " + csvFile.getName());
            file.renameTo(new File(abacRootPathDir + File.separator + _abacProcessedDir + File.separator + file.getName()));
          }
        }

      }catch (IOException ex){
        LOG.error("IO Failure ", ex);
        System.out.println("Could not process file " + file.getName());
      }
    }
  }

  private List<Sheet> getWorksheets(Workbook workbook) {
    List<Sheet> sheets = new ArrayList<>();
    for (int i = _excelSheetRange[0]; i < workbook.getNumberOfSheets() && i <= _excelSheetRange[1]; ++i) {
      sheets.add(workbook.getSheetAt(i));
    }
    return sheets;
  }

  private List<List<String>> convertToCSV(Sheet sheet) {
    List<List<String>> csvData = new ArrayList<>();
    if (sheet.getPhysicalNumberOfRows() > 0) {
      int lastRowNum = sheet.getLastRowNum();

      for (int j = _rowStart; j <= lastRowNum; ++j) {
        Row row = sheet.getRow(j);
        this.rowToCSV(csvData, row);
      }
    }

    return csvData;
  }

  private List<List<String>> convertToCSV(Workbook workbook) {

    System.out.println("Converting files contents to CSV format.");
    List<List<String>> csvData = new ArrayList<>();

    for (int i = _excelSheetRange[0]; i < workbook.getNumberOfSheets() && i <= _excelSheetRange[1]; ++i) {
      Sheet sheet = workbook.getSheetAt(i);
      if (sheet.getPhysicalNumberOfRows() > 0) {
        int lastRowNum = sheet.getLastRowNum();

        for (int j = _rowStart; j <= lastRowNum; ++j) {
          Row row = sheet.getRow(j);
          this.rowToCSV(csvData, row);
        }
      }
    }
    return csvData;
  }


  private void rowToCSV(List<List<String>> csvData, Row row) {
    ArrayList<String> csvLine = new ArrayList<>();
    if (row != null) {
      int lastCellNum = row.getLastCellNum();

      for (int i = _columnRange[0]; i < lastCellNum && i <= _columnRange[1]; ++i) {
        Cell cell = row.getCell(i);
        if (cell == null) {
          csvLine.add("");
        } else if (cell.getCellType() != 2 ) {
          csvLine.add(formatter.formatCellValue(cell));
        } else {
          csvLine.add(formatter.formatCellValue(cell, this.evaluator));
        }
      }
    }

    csvData.add(csvLine);
  }

  private Workbook openWorkbook(File file) throws IOException {
    System.out.println("Opening workbook [{}] " + file.getName());

    try (FileInputStream fis = new FileInputStream(file)) {
      return WorkbookFactory.create(fis);
    }catch (InvalidFormatException ex){
      return null;
    }
  }

  public List<File> getExcelFilesForTransform(){
    List<File> excelFiles = new ArrayList<>();
    for (String filename : getFilenames()) {
      excelFiles.add(new File(abacRootPathDir + File.separator + filename));
    }
    return excelFiles;
  }

  public String[] getFilenames(){
    DirectoryScanner scanner = new DirectoryScanner();
    scanner.setIncludes(_excelFilePatterns);
    scanner.setBasedir(abacRootPathDir);
    scanner.setCaseSensitive(false);
    scanner.scan();
    return scanner.getIncludedFiles();
  }

  private void saveCSVFile(File file, List<List<String>> csvData) throws IOException {
    System.out.println("Saving the CSV file [{}] "+ file.getName());

    try (BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file)))) {

      for (int i = 0; i < csvData.size(); ++i) {
        StringBuilder buffer = new StringBuilder();
        List<String> line = csvData.get(i);
        for (int j = 0; j < line.size(); ++j) {
           String csvLineElement  = line.get(j);
           if (csvLineElement != null) {
             buffer.append(this.escapeEmbeddedCharacters(csvLineElement)).append(
                 (line.size()-1== j ? "" : _csvSeparator) );
           }else if(line.size()-1!= j){
             buffer.append(_csvSeparator);
           }
        }
        bw.write(buffer.toString().trim());
        if (i < csvData.size() - 1) {
          bw.newLine();
        }
      }
    }
    LOG.info("Checking destination copying dirs for file " + file.getName());
    if(_abacFullPathCopyDirs != null){
      for (String abacCopyDir : _abacFullPathCopyDirs) {

        File copied = new File(abacCopyDir + File.separator + file.getName());
        LOG.info("Processing destination copying dir " + copied);
        try (
            InputStream in = new BufferedInputStream(new FileInputStream(file));
            OutputStream out = new BufferedOutputStream( new FileOutputStream(copied))) {
          LOG.info("Streams created for " + copied);
          byte[] buffer = new byte[1024];
          int lengthRead;
          while ((lengthRead = in.read(buffer)) > 0) {
            out.write(buffer, 0, lengthRead);
            out.flush();
          }
        }
      }
    }
  }

  private String escapeEmbeddedCharacters(String field) {
    if (field.contains(_csvSeparator)) {
      field = field.replaceAll(_csvSeparator, SEPARATOR_REPLACEMENT);
    }
    return field.trim();
  }


  private void validate(){
    try{
      this.abacRootPathDir = System.getProperty("abacRootPathDir");
      if(this.abacRootPathDir == null){
        throw new ExcelToCSVGenericException("abacRootPathDir not set.\n\r" + getRunningConfigString(), new Throwable());
      }
      this.abacRelativeTransformDir = System.getProperty("abacRelativeTransformDir");
      if(this.abacRelativeTransformDir == null){
        throw new ExcelToCSVGenericException("abacRelativeTransformDir not set.\n\r" + getRunningConfigString(), new Throwable());
      }else{
        File csvFolder = new File(this.abacRootPathDir + File.separator + abacRelativeTransformDir);
        if(!csvFolder.exists()){
          csvFolder.mkdir();
          if(!csvFolder.exists()){
            throw new ExcelToCSVGenericException("Processed dir "+ csvFolder.getAbsolutePath() +" could be found/create.\n\r" + getRunningConfigString(), new Throwable());
          }
        }
      }

      this.abacRelativeProcessedDir = System.getProperty("abacRelativeProcessedDir");
      if(this.abacRelativeProcessedDir != null){
        this._abacProcessedDir = this.abacRelativeProcessedDir;
      }
        File processedDir = new File(this.abacRootPathDir + File.separator + _abacProcessedDir);
        if(!processedDir.exists()){
          processedDir.mkdir();
          if(!processedDir.exists()){
            throw new ExcelToCSVGenericException("Processed dir "+ processedDir.getAbsolutePath() +" could be found/create.\n\r" + getRunningConfigString(), new Throwable());
          }
        }


      this.abacFullPathCopyDirs = System.getProperty("abacFullPathCopyDirs");
      if(this.abacFullPathCopyDirs != null){
        this._abacFullPathCopyDirs = this.abacFullPathCopyDirs.split(",");
      }

      this.excelFilePatterns = System.getProperty("excelFilePatterns");
      if(this.excelFilePatterns != null){
        this._excelFilePatterns = this.excelFilePatterns.split(",");
      }

      this.csvSeparator = System.getProperty("csvSeparator");
      if(this.csvSeparator != null){
        this._csvSeparator = this.csvSeparator;
      }

      this.excelSheetRange = System.getProperty("excelRangeTab");
      if(this.excelSheetRange != null){
        String[] excelRangeTabVals = this.excelSheetRange.split(",");
        if(excelRangeTabVals.length == 1){
          this._excelSheetRange = new int[]{Integer.valueOf(excelRangeTabVals[0]), Integer.valueOf(excelRangeTabVals[0])};
        }else {
          this._excelSheetRange = new int[]{Integer.valueOf(excelRangeTabVals[0]), Integer.valueOf(excelRangeTabVals[1])};
        }
      }

      this.rowStart = System.getProperty("rowStart");
      if(this.rowStart != null){
        this._rowStart = Integer.valueOf(this.rowStart);
      }

      this.columnRange = System.getProperty("columnRange");
      if(this.columnRange != null){
        String[] columnRangeVals = this.columnRange.split(",");
        if(columnRangeVals.length == 1){
          this._columnRange = new int[]{Integer.valueOf(columnRangeVals[0]), Integer.valueOf(this._columnRange[1])};
        }else {
          this._columnRange = new int[]{Integer.valueOf(columnRangeVals[0]), Integer.valueOf(columnRangeVals[1])};
        }
      }

      this._generateFilePerTab = System.getProperty("generateFilePerTab");
      if(this._generateFilePerTab != null && this._generateFilePerTab.equals("true")){
        this.generateFilePerTab = true;
      }
    }catch (RuntimeException ex){
      if(ex instanceof ExcelToCSVGenericException)
        throw ex;
      else{
        throw new ExcelToCSVGenericException("Invalid Configuration. Please check your java arg parameters.\n\r" + getRunningConfigString(), new Throwable());
      }
    }

  }


  private String getRunningConfigString(){
    return "Run the application using the mandatory and optional with the [] command line\n\r java "
        + "-DabacRootPathDir "
        + "-DabacRelativeTransformDir "
        + "-abacRelativeProcessedDir "
        + "[-DabacFullPathCopyDirs] "
        + "[-DexcelFilePatterns] "
        + "[-DcsvSeparator] "
        + "[-DexcelRangeTab] "
        + "[-DrowStart] "
        + "[-DcolumnRange] "
        + "[-DgenerateFilePerTab] "
        + "jar exceltocsv.jar";
  }
}
