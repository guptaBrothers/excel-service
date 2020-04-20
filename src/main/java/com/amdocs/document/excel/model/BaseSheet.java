package com.amdocs.document.excel.model;

import java.util.List;
import java.util.Map;

import lombok.Getter;
import lombok.Setter;


@Getter
@Setter
public class BaseSheet
{
    private String sheetName;
    private List<Map<Integer , Object>> rows;
}
