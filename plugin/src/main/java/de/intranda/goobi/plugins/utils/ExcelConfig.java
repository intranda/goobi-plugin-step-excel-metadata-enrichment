package de.intranda.goobi.plugins.utils;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.SubnodeConfiguration;

import lombok.Data;

@Data
public class ExcelConfig {

    private int firstLine;

    private String docstructIdentifier;

    private int rowHeader;
    private int rowDataStart;
    private int rowDataEnd;
    private List<MetadataMappingObject> metadataList = new ArrayList<>();
    private List<PersonMappingObject> personList = new ArrayList<>();
    private List<GroupMappingObject> groupList = new ArrayList<>();
    private String excelIdentifierColumn;

    /**
     * loads the &lt;config&gt; block from xml file
     * 
     * @param xmlConfig
     */

    public ExcelConfig(SubnodeConfiguration xmlConfig) {

        firstLine = xmlConfig.getInt("/firstLine", 1);
        docstructIdentifier = xmlConfig.getString("/docstructIdentifier", null);


        excelIdentifierColumn = xmlConfig.getString("/excelIdentifierColumn", null);


        rowHeader = xmlConfig.getInt("/rowHeader", 1);
        rowDataStart = xmlConfig.getInt("/rowDataStart", 2);
        rowDataEnd = xmlConfig.getInt("/rowDataEnd", 20000);

        List<HierarchicalConfiguration> mml = xmlConfig.configurationsAt("//metadata");
        for (HierarchicalConfiguration md : mml) {
            metadataList.add(getMetadata(md));
        }

        List<HierarchicalConfiguration> pml = xmlConfig.configurationsAt("//person");
        for (HierarchicalConfiguration md : pml) {
            personList.add(getPersons(md));
        }

        List<HierarchicalConfiguration> gml = xmlConfig.configurationsAt("//group");
        for (HierarchicalConfiguration md : gml) {
            String rulesetName = md.getString("@ugh");
            GroupMappingObject grp = new GroupMappingObject();
            grp.setRulesetName(rulesetName);

            String docType = md.getString("@docType", "child");
            grp.setDocType(docType);
            List<HierarchicalConfiguration> subList = md.configurationsAt("//person");
            for (HierarchicalConfiguration sub : subList) {
                PersonMappingObject pmo = getPersons(sub);
                grp.getPersonList().add(pmo);
            }

            subList = md.configurationsAt("//metadata");
            for (HierarchicalConfiguration sub : subList) {
                MetadataMappingObject pmo = getMetadata(sub);
                grp.getMetadataList().add(pmo);
            }

            groupList.add(grp);

        }
    }

    private MetadataMappingObject getMetadata(HierarchicalConfiguration md) {
        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setExcelColumn(md.getInteger("@column", null));
        mmo.setRulesetName(md.getString("@ugh"));
        mmo.setHeaderName(md.getString("@headerName", null));
        mmo.setNormdataHeaderName(md.getString("@normdataHeaderName", null));
        return mmo;
    }

    private PersonMappingObject getPersons(HierarchicalConfiguration md) {

        PersonMappingObject pmo = new PersonMappingObject();
        pmo.setRulesetName(md.getString("@ugh"));

        pmo.setFirstnameColumn(md.getInteger("firstname", null));
        pmo.setLastnameColumn(md.getInteger("lastname", null));
        pmo.setIdentifierColumn(md.getInteger("identifier", null));
        pmo.setHeaderName(md.getString("nameFieldHeader", null));
        pmo.setNormdataHeaderName(md.getString("@normdataHeaderName", null));

        pmo.setFirstnameHeaderName(md.getString("firstnameFieldHeader", null));
        pmo.setLastnameHeaderName(md.getString("lastnameFieldHeader", null));
        pmo.setSplitChar(md.getString("splitChar", " "));
        pmo.setSplitName(md.getBoolean("splitName", false));
        pmo.setFirstNameIsFirst(md.getBoolean("splitName/@firstNameIsFirstPart", false));
        return pmo;

    }

}
