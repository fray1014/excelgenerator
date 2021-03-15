import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.examples.CellStyleDetails;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellAlignment;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.*;

public class Dealer {
    //标题
    private static String title;
    //难度
    private static String level;
    //时间，单位分钟
    private static int time;
    //分量
    private static String num;
    //标签
    private static String label;
    //搭配
    private static String dapei;
    //tips
    private static String tips;
    //食材
    private static LinkedHashMap<String,String> mat;

    @Test
    public void test2(){

    }
    public void moveDir(String allPath,String path){
        File[] innerFiles = getFileArray(allPath+"\\"+path+"\\"+path);
        if(innerFiles!=null){
            for(File tmpFile : innerFiles){
                tmpFile.renameTo(new File(allPath+"\\"+path+"\\"+tmpFile.getName()));
            }
        }
        File dFile = new File(allPath+"\\"+path+"\\"+path);
        if(dFile.exists()){
            dFile.delete();
        }
    }
    @Test
    public void test() throws Exception {
        String allPath = "3.8-3\\3.8-3.14（59道）\\米博\\米博菜谱part3";
        File[] files = getFileArray(allPath);
        for(File f : files){
            if(f.isDirectory()){
                String path = f.getName().trim();
                //moveDir(allPath,path);
                String dname = allPath+"\\"+path+"\\"+path+"2.docx";
                String nname = allPath+"\\"+path+"\\"+path+"\\"+path+".docx";
                String[] text = getText(new File(dname),"");
                setProperties(text);
                writeExcel("标准化菜谱模板_FC2.xlsm",allPath,path);
            }
        }
    }
    public static String[] getText(File filePath, String nname) {
        String text = "";
        String[] res;
        String fileName = filePath.getName().toLowerCase();// 得到名字小写
        try {
            FileInputStream in = new FileInputStream(filePath);
            if (fileName.endsWith(".doc")) { // doc为后缀的

                WordExtractor extractor = new WordExtractor(in);
                text = extractor.getText();
            }
            if (fileName.endsWith(".docx")) { // docx为后缀的

                XWPFWordExtractor docx = new XWPFWordExtractor(new XWPFDocument(in));
                text = docx.getText();
            }
            in.close();
            if(!nname.equals("")){
                filePath.renameTo(new File(nname));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        res = text.split("\\n");
        return Arrays.stream(res).filter(x->x.length()!=0).toArray(String[]::new);
    }

    public static void writeExcel(String filePath,String allPath,String inName) throws IOException{
        FileInputStream in = new FileInputStream(filePath);
        XSSFWorkbook wb = new XSSFWorkbook(in);

        CellStyle cellStyle2 = wb.createCellStyle();
        //基本信息
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(1);
        row.getCell(3).setCellValue(title);

        row = sheet.getRow(2);
        row.getCell(1).setCellValue(time);
        row.getCell(3).setCellValue(level);
        row.getCell(5).setCellValue(num);

        row = sheet.getRow(4);
        row.getCell(1).setCellValue(label);

        row = sheet.getRow(5);
        row.getCell(1).setCellValue(dapei);
        row.getCell(4).setCellValue(tips);

        //可编辑的
        //String outPath = allPath+"\\"+inName+"\\"+inName+"\\"+inName+".xlsm";
        String outPath = allPath+"\\"+inName+"\\"+inName+".xlsm";
        FileOutputStream out =  new FileOutputStream(outPath);
        wb.write(out);
        out.flush();
        out.close();

        int matIndex = 8;
        for(Map.Entry<String,String> entry:mat.entrySet()){
            if(matIndex == 8){
                row = sheet.getRow(matIndex);
                //row.getCell(4).setCellValue(entry.getKey());
                row.getCell(4).setCellValue(entry.getKey().substring(0,
                        !entry.getKey().contains("（") ? entry.getKey().length():entry.getKey().indexOf("（")));
                row.getCell(5).setCellValue(yongliang(entry.getValue()));
                Cell cell = row.getCell(6);
                //cell.setCellValue(beizhu(entry.getValue()));
                cell.setCellValue(beizhu(entry.getKey()));
                cellStyle2 = row.getCell(4).getCellStyle();
                cell.setCellStyle(cellStyle2);
            }else{
                row = sheet.createRow(matIndex);
                row.setHeight(sheet.getRow(8).getHeight());
                Cell cell = row.createCell(3);
                cell.setCellStyle(cellStyle2);
                cell = row.createCell(4);
                //cell.setCellValue(entry.getKey());
                cell.setCellValue(entry.getKey().substring(0,
                        !entry.getKey().contains("（") ? entry.getKey().length():entry.getKey().indexOf("（")));
                cell.setCellStyle(cellStyle2);
                cell = row.createCell(5);
                cell.setCellValue(yongliang(entry.getValue()));
                cell.setCellStyle(cellStyle2);
                cell = row.createCell(6);
                cell.setCellValue(beizhu(entry.getValue()));
                //cell.setCellValue(beizhu(entry.getKey()));
                cell.setCellStyle(cellStyle2);
            }
            matIndex++;
        }
        //合并
        sheet.addMergedRegion(new CellRangeAddress(8, matIndex-1, 3, 3));

        //outPath = allPath+"\\"+inName+"\\"+inName+"\\"+inName+"-参考.xlsm";
        outPath = allPath+"\\"+inName+"\\"+inName+"-参考.xlsm";
        FileOutputStream out2 =  new FileOutputStream(outPath);
        wb.write(out2);
        out2.flush();
        out2.close();
        in.close();
    }

    public static void setProperties(String[] text){
        title = text[0];
        int index = findIndexPos(text,"标签：");
        String[] slabel = text[index].split("：");
        label = slabel[1].trim();
        index = findIndexPos(text,"难度：");
        String[] slevel = text[index].split("：");
        level = slevel[1].trim();
        index = findIndexPos(text,"时间：");
        time = calcTime(text[index]);
        String[] snum = text[4].split("：");
        num = snum[1].trim();
        index = findIndex(text,"搭配：");
        if(index!=0){
            String[] sdapei = text[index].split("：");
            if(sdapei.length<2){
                dapei = "";
            }else{
                dapei = sdapei[1].trim();
            }
        }else{
            dapei = "";
        }
        index = findIndex(text,"Tips");
        //判断是否有tips
        if(index != 0){
            try {
                StringBuilder sb = new StringBuilder();
                for(int i = index+1 ;i < text.length;i++){
                    if(!text[i].contains((i - index) +"、")){
                        sb.append((i-index)+"、"+text[i]+"\n");
                    }else{
                        sb.append(text[i]+"\n");
                    }
                }
                //去除多余回车
                sb.delete(sb.length() - 1, sb.length());
                tips = sb.toString();
            } catch (StringIndexOutOfBoundsException e){
                tips = "";
            }
        }else{
            tips = "";
        }
        //设置食材
        index = findIndex(text,"食材：");
        mat = new LinkedHashMap<>();
        for(int i = index+1;i<text.length;i++){
            if(text[i].length()<=2||text[i].contains("步骤：")){
                break;
            }
            //若食材用表格形式呈现
            if(text[i].contains("\\t")){
                try {
                    String[] t = text[i].split("\\t");
                    mat.put(t[0], t[1]);
                }catch (Exception e){
                    System.out.println(title);
                }
            }else{
                String[] str0 = text[i].split("\\d");
                int start = str0[0].length();
                //用量为适量
                if(start == text[i].length()){
                    String[] str1 = text[i].split(" ");
                    mat.put(str1[0],"适量");
                    //mat.put(str1[0].substring(0,str1[0].length()-2),"适量");
                }else{//用量有数字
                    //判断空格
                    if(str0[0].endsWith(" ")){
                        str0[0] = str0[0].replace(" ","");
                    }
                    mat.put(str0[0],text[i].substring(start));
                }


            }

        }
    }

    public static int firstNum(String text){
        return text.split("\\d")[0].length();
    }

    public static int calcTime(String text){
        int res = 0;
        String[] stime = text.split("：");
        String st = stime[1].trim();
        String[] hm = st.split("小时");
        try {
            //可能是X小时Y分钟
            if (text.contains("分钟")) {
                if (hm.length == 2) {
                    res = Integer.parseInt(hm[0]) * 60 + Integer.parseInt(hm[1].substring(0, hm[1].length()-"分钟".length()));
                } else {
                    res = Integer.parseInt(hm[0].substring(0, hm[0].length()-"分钟".length()));
                }
            } else {
                res = Integer.parseInt(hm[0]) * 60;
            }
        }catch (NumberFormatException e){
            System.out.println(title);
        }
        return res;
    }

    public static int findIndex(String[] text,String reg){
        int index = 0;
        for(int i = text.length-1;i >= 0;i--){
            if(text[i].contains(reg)){
                index = i;
                break;
            }
        }
        return index;
    }

    public static int findIndexPos(String[] text,String reg){
        int index = 0;
        for(int i = 0;i < text.length;i++){
            if(text[i].contains(reg)){
                index = i;
                break;
            }
        }
        return index;
    }

    public static File[] getFileArray(String strPath) {
        File dir = new File(strPath);
        return dir.listFiles();
    }

    public static String yongliang(String text){
        if(text.contains("（")){
            return text.substring(0,text.indexOf('（'));
        }
        return text;
    }

    public static String beizhu(String text){
        String res = "";
        if(text.contains("（")){
            res = text.substring(text.indexOf('（'));
        }
        return res;
    }
}
