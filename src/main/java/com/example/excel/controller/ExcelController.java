package com.example.excel.controller;

import com.example.excel.bean.Person;
import com.example.excel.service.PersonService;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;

@RestController
public class ExcelController {

    @Autowired
    private PersonService personService;

    @RequestMapping("/exportExcel")
    private void excelExport(HttpServletResponse res) {
        personService.excelExport(res);
    }

    @RequestMapping("/importExcel")
    private void importExcel() {
        personService.importExcel();
    }

}
