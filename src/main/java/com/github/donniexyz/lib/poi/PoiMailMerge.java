package com.github.donniexyz.lib.poi;

import com.github.donniexyz.lib.utils.multimap.MultiMapMinimalUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

public class PoiMailMerge {

    private static final String MAILMERGE_STRING = " MERGEFIELD  ";

    public static XWPFDocument perform(Path templatePath, Map<String, Object> dataMap) throws IOException {

        XWPFDocument doc = new XWPFDocument(
                Files.newInputStream(templatePath));

        for (IBodyElement bodyElement : doc.getBodyElements()) {
            if (BodyElementType.PARAGRAPH.equals(bodyElement.getElementType())) {
                XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                processCTP(dataMap, paragraph.getCTP());
            } else if (BodyElementType.TABLE.equals(bodyElement.getElementType())) {
                XWPFTable table = (XWPFTable) bodyElement;
                table.getCTTbl().getTrList().forEach(
                        tr -> tr.getTcList().forEach(
                                tc -> tc.getPList().forEach(
                                        p -> processCTP(dataMap, p))));
            }
        }
        return doc;
    }

    private static void processCTP(Map<String, Object> dataMap, CTP ctp) {
        processFldSimpleList(dataMap, ctp.getFldSimpleList());
        processRArray(dataMap, ctp.getRArray());
    }

    private static void processFldSimpleList(Map<String, Object> dataMap, List<CTSimpleField> fieldList) {
        for (CTSimpleField ctSimpleField : fieldList) {
            String instr = ctSimpleField.getInstr();
            if (null != instr && instr.startsWith(MAILMERGE_STRING))
                try {
                    String fieldName = instr.split(" {2}")[1];
                    ctSimpleField.getRList().get(0).getTList().get(0)
                            .setStringValue(MultiMapMinimalUtils.getFromMultiMap(fieldName, dataMap, "").toString());
                } catch (RuntimeException e) {
                    // do nothing
                }
        }
    }

    private static void processRArray(Map<String, Object> dataMap, CTR[] rArray) {
        if (rArray.length <= 1) return;
        for (int i = 0; i < rArray.length - 1; i++)
            try {
                CTText[] instrTextArray = rArray[i].getInstrTextArray();
                if (instrTextArray.length == 1) {
                    String instr = instrTextArray[0].getStringValue();
                    if (instr.startsWith(MAILMERGE_STRING)) {
                        String fieldName = instr.split(" {2}")[1].trim();
                        i++;
                        if ("separate".equals(rArray[i].getFldCharArray()[0].getFldCharType().toString())) i++;
                        CTText ctText = rArray[i].getTArray()[0];
                        String ctTextString = ctText.getStringValue();
                        if (ctTextString.startsWith("«") && ctTextString.endsWith("»")) {
                            Object fromMultiMap = MultiMapMinimalUtils.getFromMultiMap(fieldName, dataMap, "");
                            ctText.setStringValue(fromMultiMap.toString());
                        }
                    }
                }
            } catch (RuntimeException e) {
                // do nothing
            }
    }

}
