package com.example.excel.service.impl;

import com.example.excel.bean.Person;
import com.example.excel.dao.PersonDao;
import com.example.excel.service.PersonService;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileItemFactory;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.time.Period;
import java.util.*;

@Service
public class PersonServiceImpl implements PersonService {

    @Autowired
    private PersonDao personDao;

    @Override
    public void excelExport(HttpServletResponse res) {
        HashMap<String, Object> param = new HashMap<String, Object>();
        LinkedHashMap<String, String> headMap = getHeadMap();
        String fileName = "人员信息" + "-" + System.currentTimeMillis();
        List<Person> list = personDao.getPersonList();
        LinkedHashMap<String, String> personHeadMap = getHeadMap();
        HSSFWorkbook workbook = exportExcel(0, "第一页", personHeadMap);
        int num = 0;
        for (Person person : list) {
            num++;
            HSSFSheet sheet = workbook.getSheetAt(0);
            HSSFRow createRow = null;
            createRow = sheet.getRow(num);
            if (null == createRow) {
                createRow = sheet.createRow(num);
            }
            HSSFCell id = createRow.createCell(0);
            HSSFCell name = createRow.createCell(1);
            HSSFCell money = createRow.createCell(2);
            if (person != null) {
                id.setCellValue(person.getId());
                name.setCellValue(person.getName());
                money.setCellValue(person.getMoney());
            }
        }
        ByteArrayOutputStream fos = null;
        byte[] retArr = null;
        OutputStream os = null;
        try {
            fos = new ByteArrayOutputStream();
            workbook.write(fos);
            retArr = fos.toByteArray();
            fileName = new String(fileName.getBytes(), "ISO-8859-1");
            os = res.getOutputStream();
            res.reset();
            res.setHeader("Content-Disposition", "attachment; filename=" + fileName + ".xls");
            res.setContentType("application/octet-stream; charset=utf-8");
            os.write(retArr);
            os.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fos.close();
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static LinkedHashMap<String, String> getHeadMap() {
        LinkedHashMap<String, String> resultExportMap = new LinkedHashMap<String, String>();
        resultExportMap.put("id", "ID");
        resultExportMap.put("name", "名字");
        resultExportMap.put("money", "工资");
        return resultExportMap;
    }

    private HSSFWorkbook exportExcel(int sheetNum, String sheetTitle, LinkedHashMap<String, String> headMap) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(sheetNum, sheetTitle);
        sheet.setDefaultColumnWidth(10);
        //创建表头
        HSSFRow rowHead = sheet.createRow((short)0);
        int k = 0;
        for(Map.Entry<String, String> entry : headMap.entrySet() ){
            HSSFCell cell = rowHead.createCell(k++);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(entry.getValue());
        }
        return workbook;
    }

    @Override
    public void importExcel() {
        // 这里没做页面上传，直接读取本地的文件
        MultipartFile file = getMulFileByPath("D:\\Desktop\\test\\testPerson.xls");
        try {
            HSSFWorkbook wb = new HSSFWorkbook(file.getInputStream());
            Sheet sheet = wb.getSheetAt(0);
            List<Person> list = new ArrayList<>();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                try {
                    Row rows = sheet.getRow(i);
                    Person person = new Person();
                    Cell idCell = rows.getCell(0);
                    if (idCell != null) {
                        person.setId((int) idCell.getNumericCellValue());
                    }
                    Cell nameCell = rows.getCell(1);
                    if (nameCell != null) {
                        person.setName(nameCell.getStringCellValue());
                    }
                    Cell moneyCell = rows.getCell(2);
                    if (moneyCell != null) {
                        person.setMoney((long) moneyCell.getNumericCellValue());
                    }
                    list.add(person);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            personDao.insertPersonByBatch(list);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 本地文件转为MultipartFile类型(文件路径)
     */
    private static MultipartFile getMulFileByPath(String picPath) {
        FileItem fileItem = createFileItem(picPath);
        return new CommonsMultipartFile(fileItem);
    }

    private static FileItem createFileItem(String filePath)
    {
        FileItemFactory factory = new DiskFileItemFactory(16, null);
        String textFieldName = "textField";
//        int num = filePath.lastIndexOf(".");
//        String extFile = filePath.substring(num);
        FileItem item = factory.createItem(textFieldName, "text/plain", true,
                "MyFileName");
        File newfile = new File(filePath);
        long fileSize = newfile.length();
        int bytesRead = 0;
        byte[] buffer =new byte[(int) fileSize];
        try
        {
            FileInputStream fis = new FileInputStream(newfile);
            OutputStream os = item.getOutputStream();
            while ((bytesRead = fis.read(buffer, 0,  buffer.length))!= -1)
            {
                os.write(buffer, 0, bytesRead);
            }
            os.close();
            fis.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
        return item;
    }
}
