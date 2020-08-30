package excelPOI;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.xml.transform.Source;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class POIExcelReadHigh {
    /**
     * POI 读取高版本Excel文件
     *
     * @author yangtingting
     * @date 2019/07/29
     */
    public static void main(String[] args) throws Exception {
        writeExcel();
    }

    /**
     * 从excel中读取数据
     *
     * @throws IOException
     */
    public static List<Map<String, Object>> readExcel() throws IOException {
        //创建Excel，读取文件内容
        File file = new File("C:/Users/86173/Desktop/李明霞工作文档/2020考勤/2019-12月份考勤统计结果-乙方开发团队.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(FileUtils.openInputStream(file));
        //两种方式读取工作表
        // Sheet sheet=workbook.getSheet("Sheet0");
        Sheet sheet = workbook.getSheetAt(0);
        //获取sheet中最后一行行号
        int lastRowNum = sheet.getLastRowNum();
        List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();

        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            //获取当前行最后单元格列号
            int lastCellNum = row.getLastCellNum();
            Map<String, Object> map = new HashMap<String, Object>();
            for (int j = 0; j < lastCellNum; j++) {
                Cell cell = row.getCell(j);
                String value = cell.getStringCellValue();
                value = dealData(value);
                if (j == 0) {
                    map.put("name", value);
                } else if (j == 1) {
                    map.put("date", value);
                } else if (j == 2) {
                    map.put("week", value);
                } else {
                    map.put("time", value);
                }
            }
            list.add(map);
        }
        return list;
    }


    /**
     * 写数据到excel中
     *
     * @throws IOException
     */
    public static void writeExcel() throws IOException {
        //创建Excel文件薄
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建工作表sheeet
        Sheet sheet = workbook.createSheet();
        //创建第一行
        Row row = sheet.createRow(0);
        String[] title = {"姓名", "日期", "星期", "打卡时间"};
        Cell cell = null;
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
        }
        List<Map<String, Object>> list = readExcel();
        //追加数据
        for (int i = 1; i < list.size(); i++) {
            Map<String,Object> map = list.get(i);
            Row nextrow = sheet.createRow(i);
            Cell cell2 = nextrow.createCell(0);
            cell2.setCellValue(map.get("name").toString());
            cell2 = nextrow.createCell(1);
            cell2.setCellValue(map.get("date").toString());
            cell2 = nextrow.createCell(2);
            cell2.setCellValue(map.get("week").toString());
            cell2 = nextrow.createCell(3);
            cell2.setCellValue(map.get("time").toString());
        }
        //创建一个文件
        File file = new File("D:/poi_test.xlsx");
        file.createNewFile();
        FileOutputStream stream = FileUtils.openOutputStream(file);
        workbook.write(stream);
        System.out.println("写入完成");
        stream.close();
    }

    /**
     * 处理打卡时间
     *
     * @param attendanceTime
     * @return
     */
    public static String dealData(String attendanceTime) {
        String[] attendTimeArr = attendanceTime.split(" ");
        StringBuffer result = new StringBuffer();
        if (attendTimeArr.length > 3) {
            Arrays.sort(attendTimeArr);
            result.append(attendTimeArr[0]);
            result.append(" ");
            for (int i = 0; i < attendTimeArr.length; i++) {
                if (attendTimeArr[i].startsWith("12")) {
                    result.append(attendTimeArr[i]);
                    result.append(" ");
                    break;
                }
            }
            result.append(attendTimeArr[attendTimeArr.length - 1]);
            return result.toString();
        } else {
            return attendanceTime;
        }
    }
}


