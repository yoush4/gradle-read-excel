package com.accolite;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) throws IOException {
        try{
            Map<Integer, List<String>> lookup = readExcel("C:\\Users\\yousha.gharpure\\IdeaProjects\\gradle-demo1\\src\\main\\resources\\data.xlsx");
            for(Integer i : lookup.keySet()){
                List<String> str = lookup.get(i);
                System.out.println(str);
            }
        }catch (IOException e){
            System.out.println(e);
        }

    }

    public static Map<Integer, List<String>> readExcel(String filelocation) throws IOException {

        FileInputStream file = new FileInputStream(new File(filelocation));
        Workbook workbook = new XSSFWorkbook(file);



        Sheet sheet = workbook.getSheetAt(0);

        Map<Integer, List<String>> data = new HashMap<>();
        int i = 0;
        for (Row row : sheet) {
            data.put(i, new ArrayList<String>());
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case STRING:
                        data.get(i).add(cell.getRichStringCellValue().getString());
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            data.get(i)
                                    .add(cell.getDateCellValue() + "");
                        } else {
                            data.get(i).add((int)cell.getNumericCellValue() + "");
                        }
                        break;
                    case BOOLEAN:
                        //System.out.println(data.get(i).add(cell.getBooleanCellValue() + ""));
                        data.get(i).add(cell.getBooleanCellValue() + "");
                        break;
                    case FORMULA:
                        data.get(i).add(cell.getCellFormula() + "");
                        break;
                    default: data.get(i).add(" ");
                }
            }
            i++;
        }

        if (workbook != null){
            workbook.close();
        }

        return data;

    }
}