package com.amdocs.document.excel.reader;

import java.io.IOException;
import java.text.DateFormat;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public abstract class  AbstractExcelReader<T>
{
    private String sourceFile;
    private boolean parsePercentAsFloats = false;
    private int numberOfSheets = 0;
    private int rowLimit = 0; // 0 -> no limit
    private int rowOffset = 0;
    private DateFormat formatDate = null;

    public abstract List<T> covert(String fileName) throws IOException, InvalidFormatException;

    public void validate() throws Exception{

        if(sourceFile == null){
            throw new Exception("Source file name is mandatory");
        }
    }
}
