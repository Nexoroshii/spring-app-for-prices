package com.emoisei.pricesApp;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * @author Maiseichyk_YA
 */
public class WorkWithTables {
    static Map<String, double[]> ecSortedMap = new TreeMap<>();
    static Map<String, double[]> coSortedMap = new TreeMap<>();
    static Map<String, double[]> ilSortedMap = new TreeMap<>();
    static Map<String, double[]> keSingleSortedMap = new TreeMap<>();
    static Map<String, double[]> keSpraySortedMap = new TreeMap<>();
    static Map<String, double[]> keRedlandsSortedMap = new TreeMap<>();
    static Map<String, double[]> subSortedMap = new HashMap<>();

    public static void main(String[] args) throws IOException {
        Workbook wb = openFile();
        Sheet sheet = wb.getSheetAt(0);
        JsonNode json = getJsonFromSheet(sheet);
        System.out.println(json.toString());
        getAllMaps(json);
        openAndWriteToFile(wb);
    }

    public static Workbook openFile() {
        Workbook workbook;
        try {
            FileInputStream fis = new FileInputStream("fileForInput.xlsx");
            workbook = WorkbookFactory.create(fis);
            return workbook;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static JsonNode getJsonFromSheet(Sheet sheet) {
        ObjectMapper mapper = new ObjectMapper();
        ArrayNode arrayNode = mapper.createArrayNode();
        for (Row row : sheet) {
            ObjectNode objectNode = mapper.createObjectNode();
            Cell cellB = row.getCell(1); // Доступ к значению в столбце B name
            Cell cellE = row.getCell(4); // Доступ к значению в столбце E amount
            Cell cellL = row.getCell(11); // Доступ к значению в столбце L price
            Cell cellO = row.getCell(14); // Доступ к значению в столбце O label
            if (cellB != null && cellE != null && cellL != null && cellO != null) {
                String name = cellB.toString();
                String quantity = cellE.toString();
                String amount = String.valueOf((cellE.getNumericCellValue() * cellL.getNumericCellValue()));
                String label = cellO.toString();
                objectNode.put("name", name);
                objectNode.put("quantity", quantity);
                objectNode.put("amount", amount);
                objectNode.put("label", label);
                arrayNode.add(objectNode);
            }
        }
        return arrayNode;
    }

    public static void getAllMaps(JsonNode json) {
        for (JsonNode elem : json) {
            String name = elem.get("name").asText();
            double amount = elem.get("amount").asDouble();
            int quantity = elem.get("quantity").asInt();
            int label = elem.get("label").asInt();
            if (checkIfSub(name)) {
                addInMap(subSortedMap, name, amount, quantity);
            } else {
                if (label == 1) {
                    addInMap(ecSortedMap, name, amount, quantity);
                } else if (label == 2) {
                    if (name.toLowerCase().contains("престиж")) {
                        addInMap(coSortedMap, name, amount, quantity);
                    } else {
                        addInMap(ecSortedMap, name, amount, quantity);
                    }
                } else if (label == 3) {
                    if (name.toLowerCase().contains("редлендс")) {
                        addInMap(keRedlandsSortedMap, name, amount, quantity);
                    } else if (name.toLowerCase().contains("ветвистая")) {
                        addInMap(keSpraySortedMap, name, amount, quantity);
                    } else {
                        addInMap(keSingleSortedMap, name, amount, quantity);
                    }
                } else if (label == 4) {
                    addInMap(ilSortedMap, name, amount, quantity);
                } else {
                    System.out.println("no such label!!!");
                }
            }
        }
    }

    public static void addInMap(Map<String, double[]> currMap, String name, double amount, int quantity) {
        boolean found = false;
        for (String key : currMap.keySet()) {
            if (key.equalsIgnoreCase(name.toLowerCase())) {
                found=true;
                break;
            }
        }
        if (found) {
            // Если есть, то прибавить значение "amount" к существующему значению
            currMap.computeIfPresent(name, (k, acc) -> new double[]{acc[0] + amount, acc[1] + quantity});
        } else {
            // Если нет, то добавить новое имя и значение "amount" в карту
            currMap.put(name, new double[]{amount, quantity});
        }
    }

    private static boolean checkIfSub(String element) {
        //String regex2 = ".*[A-Z]{2,}";
        String regex2 = "^[A-Z]{2,}\\s.*";
        boolean isMatch = element.matches(regex2);
        boolean isMatch2 = element.contains("+");
        return isMatch || isMatch2;
    }

    private static void clearSheet(Sheet sheet) {
        // Удаляем все строки, кроме первой (оставляем заголовок)
        for (int i = sheet.getLastRowNum(); i > 0; i--) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
    }

    public static void openAndWriteToFile(Workbook workbook) {
        Sheet sheet;
        if (workbook.getSheet("New sheet5") != null) {
            sheet = workbook.getSheet("New sheet5");
            clearSheet(sheet);
        } else {
            sheet = workbook.createSheet("New sheet5");
        }
        int currRow = 0;
        Row intervalRow = sheet.createRow(currRow);
        Cell cell1 = intervalRow.createCell(0);
        cell1.setCellValue("ec");
        currRow++;
        for (String key : ecSortedMap.keySet()) {
            double[] acc = ecSortedMap.get(key);
            Row headerRow = sheet.createRow(currRow);
            Cell cell = headerRow.createCell(0);
            cell.setCellValue(key);//name
            cell = headerRow.createCell(1);
            cell.setCellValue(acc[0]);//quantity
            cell = headerRow.createCell(2);
            cell.setCellValue(acc[1]);
            double cost = Math.round((acc[0] / acc[1]) * Math.pow(10, 2)) / Math.pow(10, 2);//amount
            cell = headerRow.createCell(3);
            cell.setCellValue(cost); //cost
            cell = headerRow.createCell(4);
            cell.setCellValue(Math.round((cost * 1.55) * Math.pow(10, 2)) / Math.pow(10, 2));//finPrice
            currRow++;
        }
        intervalRow = sheet.createRow(currRow);
        cell1 = intervalRow.createCell(0);
        cell1.setCellValue("co");
        currRow++;
        for (String key : coSortedMap.keySet()) {
            double[] acc = coSortedMap.get(key);
            Row headerRow = sheet.createRow(currRow);
            Cell cell = headerRow.createCell(0);
            cell.setCellValue(key);//name
            cell = headerRow.createCell(1);
            cell.setCellValue(acc[0]);//quantity
            cell = headerRow.createCell(2);
            cell.setCellValue(acc[1]);//amount
            double cost = Math.round((acc[0] / acc[1]) * Math.pow(10, 2)) / Math.pow(10, 2);//amount
            cell = headerRow.createCell(3);
            cell.setCellValue(cost); //cost
            cell = headerRow.createCell(4);
            cell.setCellValue(Math.round((cost * 1.55) * Math.pow(10, 2)) / Math.pow(10, 2));//finPrice
            currRow++;
        }
        intervalRow = sheet.createRow(currRow);
        cell1 = intervalRow.createCell(0);
        cell1.setCellValue("ke 1");
        currRow++;
        for (String key : keSingleSortedMap.keySet()) {
            double[] acc = keSingleSortedMap.get(key);
            Row headerRow = sheet.createRow(currRow);
            Cell cell = headerRow.createCell(0);
            cell.setCellValue(key);//name
            cell = headerRow.createCell(1);
            cell.setCellValue(acc[0]);//quantity
            cell = headerRow.createCell(2);
            cell.setCellValue(acc[1]);//amount
            double cost = Math.round((acc[0] / acc[1]) * Math.pow(10, 2)) / Math.pow(10, 2);//amount
            cell = headerRow.createCell(3);
            cell.setCellValue(cost); //cost
            cell = headerRow.createCell(4);
            cell.setCellValue(Math.round((cost * 1.55) * Math.pow(10, 2)) / Math.pow(10, 2));//finPrice
            currRow++;
        }
        intervalRow = sheet.createRow(currRow);
        cell1 = intervalRow.createCell(0);
        cell1.setCellValue("ke spray");
        currRow++;
        for (String key : keSpraySortedMap.keySet()) {
            double[] acc = keSpraySortedMap.get(key);
            Row headerRow = sheet.createRow(currRow);
            Cell cell = headerRow.createCell(0);
            cell.setCellValue(key);//name
            cell = headerRow.createCell(1);
            cell.setCellValue(acc[0]);//quantity
            cell = headerRow.createCell(2);
            cell.setCellValue(acc[1]);//amount
            double cost = Math.round((acc[0] / acc[1]) * Math.pow(10, 2)) / Math.pow(10, 2);//amount
            cell = headerRow.createCell(3);
            cell.setCellValue(cost); //cost
            cell = headerRow.createCell(4);
            cell.setCellValue(Math.round((cost * 1.55) * Math.pow(10, 2)) / Math.pow(10, 2));//finPrice
            currRow++;
        }
        intervalRow = sheet.createRow(currRow);
        cell1 = intervalRow.createCell(0);
        cell1.setCellValue("ke Redlands");
        currRow++;
        for (String key : keRedlandsSortedMap.keySet()) {
            double[] acc = keRedlandsSortedMap.get(key);
            Row headerRow = sheet.createRow(currRow);
            Cell cell = headerRow.createCell(0);
            cell.setCellValue(key);//name
            cell = headerRow.createCell(1);
            cell.setCellValue(acc[0]);//quantity
            cell = headerRow.createCell(2);
            cell.setCellValue(acc[1]);//amount
            double cost = Math.round((acc[0] / acc[1]) * Math.pow(10, 2)) / Math.pow(10, 2);//amount
            cell = headerRow.createCell(3);
            cell.setCellValue(cost); //cost
            cell = headerRow.createCell(4);
            cell.setCellValue(Math.round((cost * 1.55) * Math.pow(10, 2)) / Math.pow(10, 2));//finPrice
            currRow++;
        }
        intervalRow = sheet.createRow(currRow);
        cell1 = intervalRow.createCell(0);
        cell1.setCellValue("il");
        currRow++;
        for (String key : ilSortedMap.keySet()) {
            double[] acc = ilSortedMap.get(key);
            Row headerRow = sheet.createRow(currRow);
            Cell cell = headerRow.createCell(0);
            cell.setCellValue(key);//name
            cell = headerRow.createCell(1);
            cell.setCellValue(acc[0]);//quantity
            cell = headerRow.createCell(2);
            cell.setCellValue(acc[1]);//amount
            double cost = Math.round((acc[0] / acc[1]) * Math.pow(10, 2)) / Math.pow(10, 2);//amount
            cell = headerRow.createCell(3);
            cell.setCellValue(cost); //cost
            cell = headerRow.createCell(4);
            cell.setCellValue(Math.round((cost * 1.55) * Math.pow(10, 2)) / Math.pow(10, 2));//finPrice
            currRow++;
        }
        intervalRow = sheet.createRow(currRow);
        cell1 = intervalRow.createCell(0);
        cell1.setCellValue("subs");
        currRow++;
        List<Map.Entry<String, double[]>> list = workWithSubMap(subSortedMap);
        String subName = "";
        for (Map.Entry<String, double[]> entry : list) {
            double[] acc = entry.getValue();
            Cell cell;
            if (!entry.getKey().substring(entry.getKey().lastIndexOf(' ') + 1).equals(subName)) {
                Row headerRow1 = sheet.createRow(currRow);
                subName = entry.getKey().substring(entry.getKey().lastIndexOf(' ') + 1);
                cell = headerRow1.createCell(0);
                cell.setCellValue(subName);
                currRow++;
            }
            Row headerRow = sheet.createRow(currRow);
            cell = headerRow.createCell(0);
            cell.setCellValue(entry.getKey());//name
            cell = headerRow.createCell(1);
            cell.setCellValue(acc[0]);//quantity
            cell = headerRow.createCell(2);
            cell.setCellValue(acc[1]);//amount
            double cost = Math.round((acc[0] / acc[1]) * Math.pow(10, 2)) / Math.pow(10, 2);//amount
            cell = headerRow.createCell(3);
            cell.setCellValue(cost); //cost
            currRow++;
        }
        try (FileOutputStream fileOut = new FileOutputStream("fileForInput.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static List<Map.Entry<String, double[]>> workWithSubMap(Map<String, double[]> map) {
        List<Map.Entry<String, double[]>> list = new ArrayList<>(map.entrySet());
        // Создаем компаратор для сравнения по последнему слову в ключе и алфавиту
        Comparator<Map.Entry<String, double[]>> comparator = new Comparator<Map.Entry<String, double[]>>() {
            public int compare(Map.Entry<String, double[]> entry1, Map.Entry<String, double[]> entry2) {
                // Разбиваем ключи на слова и сравниваем последние слова
                String lastWord1 = getLastWord(entry1.getKey());
                String lastWord2 = getLastWord(entry2.getKey());
                int lastWordComparison = lastWord1.compareTo(lastWord2);
                // Если последние слова не равны, возвращаем результат сравнения
                if (lastWordComparison != 0) {
                    return lastWordComparison;
                }
                // Если последние слова равны, сравниваем ключи целиком
                return entry1.getKey().compareTo(entry2.getKey());
            }

            private String getLastWord(String str) {
                // Разбиваем строку на слова и возвращаем последнее слово
                String[] words = str.split(" ");
                return words[words.length - 1];
            }
        };
        list.sort(comparator);
        return list;
    }
}
