package com.excel.report.normac;

import android.content.Context;
import android.graphics.Typeface;
import android.text.Editable;
import android.text.TextWatcher;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.AdapterView;
import android.widget.BaseExpandableListAdapter;
import android.widget.EditText;
import android.widget.Spinner;
import android.widget.TextView;

import java.util.HashMap;
import java.util.List;

/**
 * Created by JC on 2016-05-16.
 */
public class CustomExpandableListView extends BaseExpandableListAdapter {

    private Context context;
    private List<String> expandableListTitle;
    private HashMap<String, List<String>> expandableListDetail;

    public CustomExpandableListView(Context context, List<String> expandableListTitle,
                                    HashMap<String, List<String>> expandableListDetail) {
        this.context = context;
        this.expandableListTitle = expandableListTitle;
        this.expandableListDetail = expandableListDetail;
    }

    @Override
    public Object getChild(int listPosition, int expandedListPosition) {
        return this.expandableListDetail.get(this.expandableListTitle.get(listPosition))
                .get(expandedListPosition);
    }

    @Override
    public long getChildId(int listPosition, int expandedListPosition) {
        return expandedListPosition;
    }

    @Override
    public View getChildView(int listPosition, final int expandedListPosition,
                             boolean isLastChild, View convertView, ViewGroup parent) {
        final String expandedListText = (String) getChild(listPosition, expandedListPosition);
        convertView = null;
        if (convertView == null) {
            LayoutInflater layoutInflater = (LayoutInflater) this.context
                    .getSystemService(Context.LAYOUT_INFLATER_SERVICE);
            convertView = layoutInflater.inflate(R.layout.list_item, null);

            EditText inspectionDate = (EditText) convertView.findViewById(R.id.editInspectionDate);
            EditText effectiveDate = (EditText) convertView.findViewById(R.id.editEffectiveDate);

            final Spinner citySpinner = (Spinner) convertView.findViewById(R.id.spinnerCity);
            final Spinner provinceSpinner = (Spinner) convertView.findViewById(R.id.spinnerProvince);
            final Spinner appTypeSpinner = (Spinner) convertView.findViewById(R.id.spinnerAppraisalType);
            final Spinner insBySpinner = (Spinner) convertView.findViewById(R.id.spinnerInspectionBy);
            final Spinner repBySpinner = (Spinner) convertView.findViewById(R.id.spinnerReportBy);
            final Spinner asstBySpinner = (Spinner) convertView.findViewById(R.id.spinnerAssistedBy);

            final EditText cityEdit = (EditText) convertView.findViewById(R.id.editCity);
            final EditText provinceEdit = (EditText) convertView.findViewById(R.id.editProvince);
            final EditText appTypeEdit = (EditText) convertView.findViewById(R.id.editAppraisalType);
            final EditText insByEdit = (EditText) convertView.findViewById(R.id.editInspectionBy);
            final EditText repByEdit = (EditText) convertView.findViewById(R.id.editReportBy);
            final EditText asstByEdit = (EditText) convertView.findViewById(R.id.editAssistedBy);


            cityEdit.setVisibility(View.INVISIBLE);

            //EditText Listener for Inspection Date
            //Auto-appends dashes for Date
            inspectionDate.addTextChangedListener(new TextWatcher() {

                @Override
                public void beforeTextChanged(CharSequence s, int start, int count, int after) {
                }

                @Override
                public void onTextChanged(CharSequence s, int start, int before, int count) {
                }

                @Override
                public void afterTextChanged(Editable s) {

                    if (s.length() == 4 || s.length() == 7){
                        s.append('-');
                    }
                }
            });

            //EditText Listener for Effetive Date
            //Auto-appends dashes for Date
            effectiveDate.addTextChangedListener(new TextWatcher() {

                @Override
                public void beforeTextChanged(CharSequence s, int start, int count, int after) {
                }

                @Override
                public void onTextChanged(CharSequence s, int start, int before, int count) {
                }

                @Override
                public void afterTextChanged(Editable s) {

                    if (s.length() == 4 || s.length() == 7){
                        s.append('-');
                    }
                }
            });

            //Spinner Item Select Listener
            //Sets the visibility of EditText on when "other" option is selected
            citySpinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                @Override
                public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                    if (citySpinner.getSelectedItem().equals("Other")){
                        cityEdit.setVisibility(View.VISIBLE);
                    } else{
                        cityEdit.setVisibility(View.INVISIBLE);
                    }
                }

                @Override
                public void onNothingSelected(AdapterView<?> parent) {

                }
            });

            //Spinner Item Select Listener
            //Sets the visibility of EditText on when "other" option is selected
            provinceSpinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                @Override
                public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                    if (provinceSpinner.getSelectedItem().equals("Other")){
                        provinceEdit.setVisibility(View.VISIBLE);
                    } else{
                        provinceEdit.setVisibility(View.INVISIBLE);
                    }
                }

                @Override
                public void onNothingSelected(AdapterView<?> parent) {

                }
            });

            //Spinner Item Select Listener
            //Sets the visibility of EditText on when "other" option is selected
            appTypeSpinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                @Override
                public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                    if (appTypeSpinner.getSelectedItem().equals("Other")){
                        appTypeEdit.setVisibility(View.VISIBLE);
                    } else{
                        appTypeEdit.setVisibility(View.INVISIBLE);
                    }
                }

                @Override
                public void onNothingSelected(AdapterView<?> parent) {

                }
            });

            //Spinner Item Select Listener
            //Sets the visibility of EditText on when "other" option is selected
            repBySpinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                @Override
                public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                    if (repBySpinner.getSelectedItem().equals("Other")){
                        repByEdit.setVisibility(View.VISIBLE);
                    } else{
                        repByEdit.setVisibility(View.INVISIBLE);
                    }
                }

                @Override
                public void onNothingSelected(AdapterView<?> parent) {

                }
            });

            //Spinner Item Select Listener
            //Sets the visibility of EditText on when "other" option is selected
            insBySpinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                @Override
                public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                    if (insBySpinner.getSelectedItem().equals("Other")){
                        insByEdit.setVisibility(View.VISIBLE);
                    } else{
                        insByEdit.setVisibility(View.INVISIBLE);
                    }
                }

                @Override
                public void onNothingSelected(AdapterView<?> parent) {

                }
            });

            //Spinner Item Select Listener
            //Sets the visibility of EditText on when "other" option is selected
            asstBySpinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                @Override
                public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                    if (asstBySpinner.getSelectedItem().equals("Other")){
                        asstByEdit.setVisibility(View.VISIBLE);
                    } else{
                        asstByEdit.setVisibility(View.INVISIBLE);
                    }
                }

                @Override
                public void onNothingSelected(AdapterView<?> parent) {

                }
            });
        }



        //TextView expandedListTextView = (TextView) convertView
        //.findViewById(R.id.lblListItem);
        //expandedListTextView.setText(expandedListText);
        return convertView;
    }

    @Override
    public int getChildrenCount(int listPosition) {
        return this.expandableListDetail.get(this.expandableListTitle.get(listPosition))
                .size();
    }

    @Override
    public Object getGroup(int listPosition) {
        return this.expandableListTitle.get(listPosition);
    }

    @Override
    public int getGroupCount() {
        return this.expandableListTitle.size();
    }

    @Override
    public long getGroupId(int listPosition) {
        return listPosition;
    }

    @Override
    public View getGroupView(int listPosition, boolean isExpanded,
                             View convertView, ViewGroup parent) {
        String listTitle = (String) getGroup(listPosition);
        if (convertView == null) {
            LayoutInflater layoutInflater = (LayoutInflater) this.context.
                    getSystemService(Context.LAYOUT_INFLATER_SERVICE);
            convertView = layoutInflater.inflate(R.layout.list_group, null);
        }
        TextView listTitleTextView = (TextView) convertView
                .findViewById(R.id.lblListHeader);
        listTitleTextView.setTypeface(null, Typeface.BOLD);
        listTitleTextView.setText(listTitle);
        return convertView;
    }

    @Override
    public boolean hasStableIds() {
        return false;
    }

    @Override
    public boolean isChildSelectable(int listPosition, int expandedListPosition) {
        return true;
    }
}
