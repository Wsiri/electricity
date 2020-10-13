import com.alibaba.fastjson.JSON;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Collectors;

public class UsageCount {


    public static List<List<Integer>> arrList = new ArrayList<>();

    public static Map<Integer,Integer> realCountMap = new HashMap<>();

    public static Map<Integer,Integer> fakeCountMap = new HashMap<>();

    public static Integer realNum = 0;

    public static Integer fakeNum = 0;

    public static String exportPath = "/Users/wang/Desktop/ele/export/夏季工作日.xls";

    public static String readPath = exportPath.replace("export/","");


    public static void main(String[] args) throws FileNotFoundException {

        readExcel();
        String[] dataList ={"0","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23"};
        int n = 8;
        combinationSelect(dataList, n);

        List<Model> result = new ArrayList<>();
        arrList.forEach(groupList -> {

            AtomicReference<Integer> realCount = new AtomicReference<>(0);
            AtomicReference<Integer> fakeCount = new AtomicReference<>(0);

            groupList.forEach(hour ->{
                if (realCountMap.get(hour) != null ){
                    realCount.set(realCount.get() + realCountMap.get(hour));
                }
                if (fakeCountMap.get(hour) != null ){
                    fakeCount.set(fakeCount.get() + fakeCountMap.get(hour));
                }
            });


            Model model = new Model();
            model.setCount1(getPercent(realCount.get(),realNum));
            model.setCount2(getPercent(fakeCount.get(),fakeNum));
            model.setHour(groupList);
            result.add(model);
        });

        SXSSFWorkbook wb = new SXSSFWorkbook();
        //创建 Sheet页

        int total = result.size();
        int mus = 60000;

        int avg = total / mus;
        for(int i=0; i < avg + 1; i++){
            SXSSFSheet sheetA = wb.createSheet();

            SXSSFRow row = sheetA.createRow(0);
            String[] head = new String[]{"时间组合", "真正率", "假正率"};
            int headInt = 0;
            for (String title : head) {
                row.createCell(headInt++).setCellValue(title);
            }

            int num = i * mus;
            int index = 0;
            int rowInt = 1;
            for(int j=num; j<result.size(); j++){

                if (index == mus) {// 判断index == mus的时候跳出当前for循环
                    break;
                }

                //创建单元行
                row = sheetA.createRow(rowInt++);
                row.createCell(0).setCellValue(JSON.toJSONString(result.get(j).getHour()));
                row.createCell(1).setCellValue(result.get(j).getCount1());
                row.createCell(2).setCellValue(result.get(j).getCount2());

                index++;
            }

        }
        try {
            //路径需要存在
            FileOutputStream fos = new FileOutputStream(exportPath);
            wb.write(fos);
            fos.close();
            wb.close();
            System.out.println("写数据结束！");
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    /**
     * 组合选择（从列表中选择n个组合）
     * @param dataList 待选列表
     * @param n 选择个数
     */
    public static void combinationSelect(String[] dataList, int n) {
        System.out.println(String.format("C(%d, %d) = %d",
                dataList.length, n, combination(dataList.length, n)));
        combinationSelect(dataList, 0, new String[n], 0);
    }

    /**
     * 组合选择
     * @param dataList 待选列表
     * @param dataIndex 待选开始索引
     * @param resultList 前面（resultIndex-1）个的组合结果
     * @param resultIndex 选择索引，从0开始
     */
    private static void combinationSelect(String[] dataList, int dataIndex, String[] resultList, int resultIndex) {
        int resultLen = resultList.length;
        int resultCount = resultIndex + 1;
        if (resultCount > resultLen) {
            int[] intArray =  Arrays.stream(resultList).mapToInt(Integer::parseInt).toArray();
            arrList.add(Arrays.stream(intArray).boxed().collect(Collectors.toList()));
            return;
        }

        // 递归选择下一个
        for (int i = dataIndex; i < dataList.length + resultCount - resultLen; i++) {
            resultList[resultIndex] = dataList[i];
            combinationSelect(dataList, i + 1, resultList, resultIndex + 1);
        }
    }


    public static long combination(int m, int n) {
        return m <= n ? factorial(n) / (factorial(m) * factorial((n - m))) : 0;
    }

    private static long factorial(int n) {
        long sum = 1;
        while( n > 0 ) {
            sum = sum * n--;
        }
        return sum;
    }



    public static void readExcel() {
        try {
            // 创建输入流，读取Excel
            InputStream is = new FileInputStream(readPath);
            POIFSFileSystem fs = new POIFSFileSystem(is);
            HSSFWorkbook wb = new HSSFWorkbook(fs);

            String strDateFormat = "yyyy-MM-dd HH:mm:ss";
            SimpleDateFormat sdf = new SimpleDateFormat(strDateFormat);
            //遍历Sheet页
            for(int sheet=0; sheet < wb.getNumberOfSheets(); sheet++){
                HSSFSheet s = wb.getSheetAt(sheet);
                System.out.println(s.getSheetName());
                if(s == null){
                    continue;
                }
                Integer tenPercent = s.getLastRowNum() / 10;
                realNum = tenPercent;
                fakeNum = s.getLastRowNum() - realNum;
                //遍历row
                for(int row = 1; row <= s.getLastRowNum(); row++){
                    HSSFRow r = s.getRow(row);
                    if(r == null){
                        continue;
                    }
                    HSSFCell c = r.getCell(0);
                    String result = c.toString();
                    Date date = sdf.parse(result);
                    Integer hour = date.getHours();
                    if (row <= tenPercent) {
                        realCountMap.merge(hour, 1, Integer::sum);
                    } else {
                        fakeCountMap.merge(hour,1, Integer::sum);
                    }
                }
            }
            if(is != null){
                is.close();
            }
            if(wb != null){
                wb.close();
            }
        } catch (IOException | ParseException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

    private static String getPercent(Integer x, double total) {
        String resultStr= "";//接受百分比的值

        double son = (double) x;
        double mom = total;
        double tempResult = son / mom;
        NumberFormat nf  =  NumberFormat.getPercentInstance();
        nf.setMinimumFractionDigits( 2 );
        resultStr = nf.format(tempResult);
        return  resultStr;
    }
}
