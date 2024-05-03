package de.intranda.goobi.plugins;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.DirectoryStream;
import java.nio.file.Path;

/**
 * This file is part of a plugin for Goobi - a Workflow tool for the support of mass digitization.
 *
 * Visit the websites for more information.
 *          - https://goobi.io
 *          - https://www.intranda.com
 *          - https://github.com/intranda/goobi
 *
 * This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free
 * Software Foundation; either version 2 of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
 * FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59
 * Temple Place, Suite 330, Boston, MA 02111-1307 USA
 *
 */

import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.configuration.SubnodeConfiguration;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.goobi.beans.Process;
import org.goobi.beans.Step;
import org.goobi.production.enums.PluginGuiType;
import org.goobi.production.enums.PluginReturnValue;
import org.goobi.production.enums.PluginType;
import org.goobi.production.enums.StepReturnValue;
import org.goobi.production.plugin.interfaces.IStepPluginVersion2;

import de.intranda.goobi.plugins.utils.ExcelConfig;
import de.intranda.goobi.plugins.utils.MetadataMappingObject;
import de.intranda.goobi.plugins.utils.PersonMappingObject;
import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.helper.StorageProvider;
import de.sub.goobi.helper.exceptions.DAOException;
import de.sub.goobi.helper.exceptions.SwapException;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.log4j.Log4j2;
import net.xeoh.plugins.base.annotations.PluginImplementation;
import ugh.dl.DigitalDocument;
import ugh.dl.DocStruct;
import ugh.dl.Fileformat;
import ugh.dl.Metadata;
import ugh.dl.MetadataType;
import ugh.dl.Person;
import ugh.dl.Prefs;
import ugh.exceptions.MetadataTypeNotAllowedException;
import ugh.exceptions.PreferencesException;
import ugh.exceptions.ReadException;
import ugh.exceptions.WriteException;

@PluginImplementation
@Log4j2
public class ExcelMetadataenrichmentStepPlugin implements IStepPluginVersion2 {

    // TODO where does the excel file come from? File system, file upload?

    // TODO enrich existing elements or create new ones?

    @Getter
    private String title = "intranda_step_excelMetadataenrichment";
    @Getter
    private Step step;

    private String returnPath;

    private Process process;
    private Prefs prefs;
    private ExcelConfig ec;

    @Getter
    @Setter
    private String excelFile = null;

    @Override
    public void initialize(Step step, String returnPath) {
        this.returnPath = returnPath;
        this.step = step;
        process = step.getProzess();
        prefs = process.getRegelsatz().getPreferences();

        // read configuration file
        SubnodeConfiguration myconfig = ConfigPlugins.getProjectAndStepConfig(title, step);

        ec = new ExcelConfig(myconfig);

    }

    @Override
    public PluginGuiType getPluginGuiType() {
        return PluginGuiType.NONE;
    }

    @Override
    public String getPagePath() {
        return "/uii/plugin_step_excelMetadataenrichment.xhtml";
    }

    @Override
    public PluginType getType() {
        return PluginType.Step;
    }

    @Override
    public String cancel() {
        return "/uii" + returnPath;
    }

    @Override
    public String finish() {
        return "/uii" + returnPath;
    }

    @Override
    public int getInterfaceVersion() {
        return 0;
    }

    @Override
    public HashMap<String, StepReturnValue> validate() {
        return null;
    }

    @Override
    public boolean execute() {
        PluginReturnValue ret = run();
        return ret != PluginReturnValue.ERROR;
    }

    @Override
    public PluginReturnValue run() {

        Fileformat fileformat = null;
        DigitalDocument digitalDocument = null;
        DocStruct logical = null;
        try {
            // read mets file
            fileformat = process.readMetadataFile();
            digitalDocument = fileformat.getDigitalDocument();
            logical = digitalDocument.getLogicalDocStruct();
            if (logical.getType().isAnchor()) {
                logical = logical.getAllChildren().get(0);
            }

        } catch (ReadException | PreferencesException | IOException | SwapException e) {
            log.error(e);
            return PluginReturnValue.ERROR;
        }

        // its always null unless we are in a junit test
        if (excelFile == null) {
            String folder = null;
            // we have an existing folder
            if (ec.getExcelFolder().contains("/")) {
                folder = ec.getExcelFolder();
            } else {
                // we have a folder variable
                try {
                    folder = process.getConfiguredImageFolder(ec.getExcelFolder());
                } catch (IOException | SwapException | DAOException e) {
                    log.error(e);
                }
            }

            List<Path> excelFilesInFolder = StorageProvider.getInstance().listFiles(folder, EXCEL_FILTER);
            if (excelFilesInFolder.size() == 1) {
                excelFile = excelFilesInFolder.get(0).toString();
            } else {
                for (Path file : excelFilesInFolder) {
                    if (file.getFileName().toString().toLowerCase().equals(process.getTitel().toLowerCase() + ".xlsx")) {
                        excelFile = file.toString();
                        break;
                    }
                }
            }
        }
        // abort if no file was found
        if (excelFile == null) {
            log.error("No import file found for process {}", process.getId());
            return PluginReturnValue.ERROR;
        }

        // read excel file

        Map<String, Map<Integer, String>> completeMap = new HashMap<>();

        String idColumn = ec.getExcelIdentifierColumn();
        Map<String, Integer> headerOrder = new HashMap<>();
        InputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(excelFile);
            BOMInputStream in = new BOMInputStream(fileInputStream, false);
            Workbook wb = WorkbookFactory.create(in);
            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.rowIterator();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            // get header and data row number from config first
            int rowHeader = ec.getRowHeader();
            int rowDataStart = ec.getRowDataStart();
            int rowDataEnd = ec.getRowDataEnd();
            int rowCounter = 0;

            //  find the header row
            Row headerRow = null;
            while (rowCounter < rowHeader) {
                headerRow = rowIterator.next();
                rowCounter++;
            }

            //  read and validate the header row
            int numberOfCells = headerRow.getLastCellNum();
            for (int i = 0; i < numberOfCells; i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    String value = cell.getStringCellValue();
                    headerOrder.put(value, i);
                }
            }

            // find out the first data row
            while (rowCounter < rowDataStart - 1) {
                headerRow = rowIterator.next();
                rowCounter++;
            }

            while (rowIterator.hasNext() && rowCounter < rowDataEnd) {
                Map<Integer, String> rowMap = new HashMap<>();
                Row row = rowIterator.next();
                rowCounter++;
                int lastColumn = row.getLastCellNum();
                if (lastColumn == -1) {
                    continue;
                }
                for (int cn = 0; cn < lastColumn; cn++) {
                    Cell cell = row.getCell(cn, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String value = "";
                    switch (cell.getCellType()) {
                        case BOOLEAN:
                            value = cell.getBooleanCellValue() ? "true" : "false";
                            break;
                        case FORMULA:
                            CellValue cellValue = evaluator.evaluate(cell);
                            switch (cellValue.getCellType()) {
                                case NUMERIC:
                                    value = String.valueOf((long) cell.getNumericCellValue());
                                    break;
                                case STRING:
                                    value = cell.getStringCellValue();
                                    break;
                                default:
                                    value = "";
                                    break;
                            }
                            break;
                        case NUMERIC:
                            double val = cell.getNumericCellValue();
                            String stringValue = String.valueOf(val);
                            if (stringValue.endsWith(".0")) {
                                value = String.valueOf((long) cell.getNumericCellValue());
                            } else {
                                value = stringValue;
                            }
                            break;
                        case STRING:
                            value = cell.getStringCellValue();
                            break;
                        default:
                            value = "";
                            break;
                    }
                    rowMap.put(cn, value);
                }

                String identifier = rowMap.get(headerOrder.get(idColumn));
                completeMap.put(identifier, rowMap);
            }

        } catch (Exception e) {
            log.error(e);
            return PluginReturnValue.ERROR;
        } finally {
            if (fileInputStream != null) {
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    log.error(e);
                    return PluginReturnValue.ERROR;
                }
            }
        }

        // find structure element for each row
        List<DocStruct> children = logical.getAllChildrenAsFlatList();

        MetadataType identifierType = prefs.getMetadataTypeByName(ec.getDocstructIdentifier());

        for (DocStruct child : children) {
            // get identifier from docstruct
            List<? extends Metadata> md = child.getAllMetadataByType(identifierType);
            if (md != null && !md.isEmpty()) {
                // search for excel metadata with this identifier
                String docstructId = md.get(0).getValue();
                Map<Integer, String> rowMap = completeMap.get(docstructId);
                // add  metadata
                if (rowMap == null) {
                    log.info("Skip import for " + docstructId);
                    continue;
                }

                for (MetadataMappingObject mmo : ec.getMetadataList()) {

                    String metadataValue = rowMap.get(headerOrder.get(mmo.getHeaderName()));
                    String identifier = null;
                    if (mmo.getNormdataHeaderName() != null) {
                        identifier = rowMap.get(headerOrder.get(mmo.getNormdataHeaderName()));
                    }
                    MetadataType type = prefs.getMetadataTypeByName(mmo.getRulesetName());
                    // TODO remove/overwrite/skip existing fields?
                    List<? extends Metadata> mdl = child.getAllMetadataByType(type);
                    if (mdl != null && !mdl.isEmpty()) {
                        Metadata existingMetadata = mdl.get(0);
                        existingMetadata.setValue(metadataValue);
                    } else if (StringUtils.isNotBlank(metadataValue)) {
                        try {
                            Metadata metadata = new Metadata(type);
                            metadata.setValue(metadataValue);
                            if (StringUtils.isNotBlank(identifier)) {
                                metadata.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);
                            }
                            child.addMetadata(metadata);
                        } catch (MetadataTypeNotAllowedException e) {
                            // metadata is not allowed, ignore it
                        }

                    }
                }

                for (PersonMappingObject mmo : ec.getPersonList()) {
                    String firstname = "";
                    String lastname = "";
                    if (mmo.isSplitName()) {
                        String name = rowMap.get(headerOrder.get(mmo.getHeaderName()));
                        if (StringUtils.isNotBlank(name)) {
                            if (name.contains(mmo.getSplitChar())) {
                                if (mmo.isFirstNameIsFirst()) {
                                    firstname = name.substring(0, name.lastIndexOf(mmo.getSplitChar()));
                                    lastname = name.substring(name.lastIndexOf(mmo.getSplitChar()));
                                } else {
                                    lastname = name.substring(0, name.lastIndexOf(mmo.getSplitChar())).trim();
                                    firstname = name.substring(name.lastIndexOf(mmo.getSplitChar()) + 1).trim();
                                }
                            } else {
                                lastname = name;
                            }
                        }
                    } else {
                        firstname = rowMap.get(headerOrder.get(mmo.getFirstnameHeaderName()));
                        lastname = rowMap.get(headerOrder.get(mmo.getLastnameHeaderName()));
                    }

                    String identifier = null;
                    if (mmo.getNormdataHeaderName() != null) {
                        identifier = rowMap.get(headerOrder.get(mmo.getNormdataHeaderName()));
                    }
                    if (StringUtils.isNotBlank(mmo.getRulesetName())) {
                        try {
                            Person p = new Person(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                            p.setFirstname(firstname);
                            p.setLastname(lastname);

                            if (identifier != null) {
                                p.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);
                            }

                            child.addPerson(p);

                            //                            logical.addPerson(p);
                        } catch (MetadataTypeNotAllowedException e) {
                            log.info(e);
                            // Metadata is not known or not allowed
                        }
                    }
                }

            }
        }

        //  save mets file
        try {
            process.writeMetadataFile(fileformat);
        } catch (WriteException | PreferencesException | IOException | SwapException e) {
            log.error(e);
            return PluginReturnValue.ERROR;
        }

        return PluginReturnValue.FINISH;
    }

    public static final DirectoryStream.Filter<Path> EXCEL_FILTER = new DirectoryStream.Filter<Path>() {

        @Override
        public boolean accept(Path path) {
            String name = path.getFileName().toString();
            return name.toLowerCase().endsWith(".xlsx");
        }

    };
}
