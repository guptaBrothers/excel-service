package com.amdocs.document.excel.controller;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.amdocs.document.excel.service.ExcelService;

@RestController
@RequestMapping("/excel-service")
public class ExcelController
{
    @Autowired
    private ExcelService excelService;

    @RequestMapping("/convert-to-json/{excelFile}")
    public ResponseEntity<Object> fetchDataFromExcel(@PathVariable ("excelFile") String fileName)
            throws IOException, InvalidFormatException
    {

        return  new ResponseEntity<>(excelService.getData(fileName), HttpStatus.OK);
    }
}
