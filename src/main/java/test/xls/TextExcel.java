package test.xls;

import jakarta.annotation.PostConstruct;
import java.io.BufferedInputStream;

import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.Duration;
import java.time.Instant;
import java.util.LinkedList;
import java.util.List;
import java.util.stream.Stream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

@Component
public class TextExcel {
    @PostConstruct
    public void init() throws IOException {
        try {
            //IOUtils.setByteArrayMaxOverride(1_000_000_000);
            var start = Instant.now();
            FileInputStream file = new FileInputStream(new File("/Users/kirill_usov/Temp/2m_file.xls"));
            POIFSFileSystem fsPOI = new POIFSFileSystem(new BufferedInputStream(file));
       
            HSSFWorkbook workbook = new HSSFWorkbook(fsPOI);

            var sheet = workbook.getSheetAt(0);

            var list = new LinkedList<String>();
            int i = 0;
            for (var row : sheet) {
                StringBuilder sb = new StringBuilder();
                sb.append(i).append(";");
                
                for (Cell cell : row) {
                     switch (cell.getCellType()) {
                         case STRING: sb.append(cell.getStringCellValue()).append(";"); break;
                         case NUMERIC: sb.append(cell.getNumericCellValue()).append(";"); break;
                         case BOOLEAN: sb.append(cell.getBooleanCellValue()).append(";"); break;
                         case FORMULA: sb.append(cell.getCellFormula()).append(";"); break;
                         default:
                     }
                }
                
                list.add(sb.toString());

                i++;
            }
            
            var end = Instant.now();
        
            System.out.println(Duration.between(start, end));
        } catch (Exception ex) {
            int r = 0;
        }
        
        var start = Instant.now();
//        
        //List<String> list = new LinkedList<>();
        try (InputStream is = new FileInputStream(new File("/Users/kirill_usov/Temp/2m_file.xls")); ReadableWorkbook wb = new ReadableWorkbook(is)) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                rows.forEach(row -> {
                    // BigDecimal num = r.getCellAsNumber(0).orElse(null);
                    // String str = r.getCellAsString(1).orElse(null);
                    // LocalDateTime date = r.getCellAsDate(2).orElse(null);
                    // Row row = ...;
                    StringBuilder sb = new StringBuilder();
                    row.forEach(cell -> {
                        var value = cell.getRawValue();
                        sb.append(value).append(";");
                    });

                    //list.add(sb.toString());
                });
            }
        }
        
        var end = Instant.now();
        
        System.out.println(Duration.between(start, end));
    }
}
