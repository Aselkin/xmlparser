import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ReadExcelDemo {
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream(new File("B777 ATA25.xlsx"));
            FileInputStream fileata = new FileInputStream(new File("subata_multi.xlsx"));


            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFWorkbook wbsubata = new XSSFWorkbook(fileata);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet sheetata;
            Integer qtywords = 1;
            String patternType = "(\\d{2}-+\\d{2}-+\\d{2})";
            if (sheet.getRow(1).getCell(6).getStringCellValue().trim().equals("B777")) {
                sheetata = wbsubata.getSheetAt(1);
                qtywords = 0;
                patternType = "(\\d{2}-+\\d{2})";

            } else {
                sheetata = wbsubata.getSheetAt(0);
            }



            deleteColumn(sheet, 25);


            //Iterate through each rows one by on
            for (int i = 1; i < sheet.getLastRowNum(); ++i) {

                Row row = sheet.getRow(i);
                //For each row, iterate through all the columns

                Cell colE = row.getCell(4);
                Cell colY = row.getCell(24);
                Cell colX = row.getCell(23);
                Cell colL = row.getCell(11);

                //System.out.println(colL);

                if (colE == null || colY == null || colL == null || colX == null) continue;
                if (colL.getStringCellValue().equals("N")) continue;

                String cellE = colE.getStringCellValue();
                System.out.println("Origin subata:" + cellE);

                String cellYsh = subata(colY);
                if ("not found".equals(cellYsh)) {
                    for (int k = 1; k < sheetata.getLastRowNum(); ++k) {
                        Row rowata = sheetata.getRow(k);

                        Cell colA = rowata.getCell(0);
                        Cell colB = rowata.getCell(1);
                        String[] colBshort = validateCellata(colB).split(" ");
                        int qtyWord = 0;

                        String allWords = "";
                        for (String word : colBshort) {
                            if (word.length() > 3 && colY.getStringCellValue().toUpperCase().contains(word) && !allWords.contains(word)) {
                                qtyWord++;
                                allWords += word + " ";
                            }
                            else if (allWords.equals("") && word.length() > 3 && colX.getStringCellValue().toUpperCase().contains(word) && !allWords.contains(word)) {
                                qtyWord++;
                                allWords += word + " ";
                            }
                        }



                        if (qtyWord > qtywords) {
                            Pattern patternata = Pattern.compile(patternType);
                            Matcher matcherata = patternata.matcher(colA.getStringCellValue());
                            if (matcherata.find()) {
                                String ata = matcherata.group(0).substring(0, 5).replace("-", "");
                                if (!"not found".equals(cellYsh) && !cellYsh.contains(ata)) {
                                    cellYsh += ", " + ata;
                                } else {
                                    cellYsh = ata;
                                }
                                System.out.println("Subata:" + cellYsh + "\nTEMPLATE:" + colY.getStringCellValue().toUpperCase() + "\nFOUND:" + allWords);
                            }
                        }


                    }
                }
//                System.out.println(cellYsh);

                if (!cellE.trim().equals(cellYsh.trim())) {
                    Cell cell_new = row.createCell(row.getLastCellNum());
                    cell_new.setCellValue(cellYsh);
                }
            }
            file.close();

            FileOutputStream out = new FileOutputStream(new File("B777 ATA25.xlsx"));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String subata(Cell celly) {
        Pattern pattern = Pattern.compile("AMM.*(\\d{2}-+\\d{2}-+\\d{2})");
        Matcher matcher = pattern.matcher(validateCell(celly));
        if (matcher.find()) {
            return matcher.group(1).substring(0, 5).replace("-", "");
        }
        return "not found";
    }

    private static String validateCell(Cell celly) {
        return celly.getStringCellValue()
                .replaceAll("\n", "")
                .replaceAll(" ", "")
                .replaceAll(":", "");
    }

    private static String validateCellata(Cell celly) {
        return celly.getStringCellValue().toUpperCase()
                .replaceAll("\n", "")
                .replaceAll("-", "")
                .replaceAll("  ", " ")
                .replaceAll(" AND ", " ")
                .replaceAll(" TEST ", " ");
    }


    public static void deleteColumn(XSSFSheet sheettodel, int rowtodel) {
        for (int j = 0; j <= sheettodel.getLastRowNum(); j++) {
            Row row = sheettodel.getRow(j);
            if (row.getLastCellNum() >= rowtodel) {
                Cell celltodel = row.getCell(rowtodel);
                if (celltodel == null) continue;
                row.removeCell(celltodel);
            }

        }
    }


}