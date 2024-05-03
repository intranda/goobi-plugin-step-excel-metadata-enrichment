package de.intranda.goobi.plugins.utils;


import lombok.Data;

@Data
public class MetadataMappingObject {

    private String rulesetName;
    private Integer excelColumn;
    private String headerName;

    private String normdataHeaderName;

}
