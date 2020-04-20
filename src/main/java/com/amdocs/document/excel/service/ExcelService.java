package com.amdocs.document.excel.service;

import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.amdocs.document.excel.model.UserSheet;
import com.amdocs.document.excel.reader.ExcelFileReader;

@Service
public class ExcelService
{

    @Autowired
    private ExcelFileReader excelReader;

    public List<UserSheet> getData(String fileName) throws IOException, InvalidFormatException
    {

        return excelReader.covert(fileName);
    }
}
