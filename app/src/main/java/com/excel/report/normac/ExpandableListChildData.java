package com.excel.report.normac;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * Created by JC on 2016-05-16.
 */
public class ExpandableListChildData {
    public static HashMap<String, List<String>> getData() {
        HashMap<String, List<String>> expandableListDetail = new HashMap<String, List<String>>();

        // Adding child data to General Report Information
        List<String> generalInfo = new ArrayList<String>();
        generalInfo.add("Report:");


        expandableListDetail.put("Field Notes Form", generalInfo);


        return expandableListDetail;
    }
}
