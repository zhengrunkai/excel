package com.example.excel.service;

import javax.servlet.http.HttpServletResponse;

public interface PersonService {

    void excelExport(HttpServletResponse res);

    void importExcel();

}
