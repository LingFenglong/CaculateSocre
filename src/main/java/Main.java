import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.StreamSupport;

public class Main {
    private static List<String> title;
    private static List<List<String>> body;
    private static int scoreIndex;
    private static int GPAIndex;
    private static int creditIndex;
    private static int courseNatureIndex;

    public static void main(String[] args) throws IOException, InvalidFormatException {
        if (args == null || args.length == 0) {
            System.out.println("未指定参数");
            System.exit(1);
        }
        XSSFWorkbook sheets = new XSSFWorkbook(new File(args[0]));
        XSSFSheet sheet = sheets.getSheetAt(0);

        title = getTitle(sheet);
        body = getBody(sheet);
        scoreIndex = getScoreIndex();
        creditIndex = getCreditIndex();
        GPAIndex = getGPAIndex();
        courseNatureIndex = getCourseNatureIndex();

        Double a =  calculateA();   // 计算学分加权平均分
        Double b = calculateB();   // 计算平均学分绩点

        System.out.printf("学分加权平均分：%.2f\r\n平均学分绩点：%.2f%n", a, b);

        sheets.close();
    }

    private static int getCourseNatureIndex() {
        return title.indexOf("课程性质");
    }

    private static int getCreditIndex() {
        return title.indexOf("学分");
    }

    private static int getGPAIndex() {
        return title.indexOf("绩点");
    }

    private static int getScoreIndex() {
        return title.indexOf("成绩");
    }

    private static List<List<String>> getBody(XSSFSheet sheet) {
        return StreamSupport
                .stream(sheet.spliterator(), true)
                .map(cells -> StreamSupport
                        .stream(cells.spliterator(),false)
                        .map(Cell::getStringCellValue)
                        .toList())
                .toList();
    }


    private static List<String> getTitle(XSSFSheet sheet) {
        return StreamSupport
                .stream(sheet.getRow(0).spliterator(), false)
                .map(Cell::getStringCellValue)
                .toList();
    }

    private static Double calculateA() {
        List<Double> result = body.stream()
                .filter(row -> row.get(courseNatureIndex).equals("必修"))
                .map(row -> List.of(
                        Double.parseDouble(row.get(creditIndex)),
                        Double.parseDouble(row.get(creditIndex)) * (row.get(scoreIndex).equals("合格") ? 75 : Double.parseDouble(row.get(scoreIndex)))))
                .reduce((older, newer) -> List.of(older.get(0) + newer.get(0), older.get(1) + newer.get(1)))
                .orElseGet(ArrayList::new);
        return result.get(1) / result.get(0);
    }

    private static Double calculateB() {
        List<Double> result = body.stream()
                .filter(row -> row.get(courseNatureIndex).equals("必修"))
                .map(row -> List.of(
                        Double.parseDouble(row.get(creditIndex)),
                        Double.parseDouble(row.get(creditIndex)) * Double.parseDouble(row.get(GPAIndex))))
                .reduce((older, newer) -> List.of(older.get(0) + newer.get(0), older.get(1) + newer.get(1)))
                .orElseGet(ArrayList::new);
        return result.get(1) / result.get(0);
    }
}