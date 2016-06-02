/* MainActivity.java
Description:

*/
package com.excel.report.normac;

import android.app.Activity;
import android.content.Context;
import android.content.DialogInterface;
import android.graphics.Bitmap;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.os.Handler;
import android.provider.MediaStore;
import android.support.design.widget.NavigationView;
import android.support.v4.view.GravityCompat;
import android.support.v4.widget.DrawerLayout;
import android.support.v7.app.AlertDialog;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.AdapterView;
import android.widget.Button;
import android.widget.CheckBox;
import android.widget.ExpandableListAdapter;
import android.widget.ExpandableListView;
import android.widget.GridView;
import android.widget.ImageView;
import android.widget.RadioButton;
import android.widget.Spinner;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;


import android.support.design.widget.FloatingActionButton;
import android.support.v7.app.ActionBarDrawerToggle;
import android.support.v7.widget.Toolbar;
import android.content.Intent;

import android.widget.EditText;
import android.widget.ViewSwitcher;

import com.nostra13.universalimageloader.cache.memory.impl.WeakMemoryCache;
import com.nostra13.universalimageloader.core.DisplayImageOptions;
import com.nostra13.universalimageloader.core.ImageLoader;
import com.nostra13.universalimageloader.core.ImageLoaderConfiguration;
import com.nostra13.universalimageloader.core.assist.ImageScaleType;

public class MainActivity extends AppCompatActivity implements NavigationView.OnNavigationItemSelectedListener {


    static String TAG = "ExelLog";
    ExpandableListView expandableListView;
    ExpandableListAdapter expandableListAdapter;
    List<String> expandableListTitle;
    HashMap<String, List<String>> expandableListDetail;
    private static final int CAPTURE_IMAGE_ACTIVITY_REQUEST_CODE = 100;
    private Uri fileUri;
    public static final int MEDIA_TYPE_IMAGE = 1;


    GridView gridGalleryy;
    Handler handler;
    CustomGalleryAdapter adapter;
    ArrayList<String> imagePaths;
    ImageView imgSinglePick;
    Button btnGalleryPick;
    Button btnGalleryPickMul;

    String action;
    ViewSwitcher viewSwitcher;
    ImageLoader imageLoader;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        this.setTitle(getResources().getString(R.string.app_long_name));

        expandableListView = (ExpandableListView) findViewById(R.id.expandableListView);
        expandableListDetail = ExpandableListChildData.getData();

        expandableListTitle = new ArrayList<String>(expandableListDetail.keySet());
        Collections.sort(expandableListTitle);

        expandableListAdapter = new CustomExpandableListView(this, expandableListTitle, expandableListDetail);
        expandableListView.setAdapter(expandableListAdapter);

        //////////////
        // Disables the expansion on a group item once click
// This prevents data loss when you collapse the group item
        expandableListView.setOnGroupClickListener(new ExpandableListView.OnGroupClickListener() {
            @Override
            public boolean onGroupClick(ExpandableListView parent, View v, int groupPosition, long id) {
                return parent.isGroupExpanded(groupPosition);
            }
        });
////////////////

        expandableListView.setOnGroupExpandListener(new ExpandableListView.OnGroupExpandListener() {

            @Override
            public void onGroupExpand(int groupPosition) {
                Toast.makeText(getApplicationContext(),
                        expandableListTitle.get(groupPosition) + " List Expanded.",
                        Toast.LENGTH_SHORT).show();
            }
        });

        expandableListView.setOnGroupCollapseListener(new ExpandableListView.OnGroupCollapseListener() {

            @Override
            public void onGroupCollapse(int groupPosition) {
                Toast.makeText(getApplicationContext(),
                        expandableListTitle.get(groupPosition) + " List Collapsed.",
                        Toast.LENGTH_SHORT).show();

            }
        });

        expandableListView.setOnChildClickListener(new ExpandableListView.OnChildClickListener() {
            @Override
            public boolean onChildClick(ExpandableListView parent, View v,
                                        int groupPosition, int childPosition, long id) {
                Toast.makeText(
                        getApplicationContext(),
                        expandableListTitle.get(groupPosition)
                                + " -> "
                                + expandableListDetail.get(
                                expandableListTitle.get(groupPosition)).get(
                                childPosition), Toast.LENGTH_SHORT
                ).show();
                return false;
            }
        });

        Toolbar toolbar = (Toolbar) findViewById(R.id.toolbar);
        setSupportActionBar(toolbar);
////////////////////// CAMERA


        FloatingActionButton fab = (FloatingActionButton) findViewById(R.id.fab);
        fab.setOnClickListener(new View.OnClickListener() {

            @Override
            public void onClick(View view) {


                /*
                Intent intent = new Intent(android.provider.MediaStore.ACTION_IMAGE_CAPTURE);



                File picDirectory = new File(String.valueOf(Environment.getExternalStoragePublicDirectory(
                       Environment.DIRECTORY_PICTURES + "/Report_Images/" + "Images_" + formattedDate )));
                picDirectory.mkdirs();

               File image = new File(picDirectory, editStrata.getText().toString() + "_" + editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() +  "_" + formattedDate.toString() + ".jpg");
                Uri uriSavedImage = Uri.fromFile(image);


                intent.putExtra(MediaStore.EXTRA_OUTPUT, uriSavedImage);


                startActivityForResult(intent, CAPTURE_IMAGE_ACTIVITY_REQUEST_CODE);*/

                Intent intent = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);

                fileUri = getOutputMediaFileUri(MEDIA_TYPE_IMAGE); // create a file to save the image
                intent.putExtra(MediaStore.EXTRA_OUTPUT, fileUri); // set the image file name

                // start the image capture Intent
                startActivityForResult(intent, CAPTURE_IMAGE_ACTIVITY_REQUEST_CODE);


            }
        });

        DrawerLayout drawer = (DrawerLayout) findViewById(R.id.drawer_layout);
        ActionBarDrawerToggle toggle = new ActionBarDrawerToggle(
                this, drawer, toolbar, R.string.navigation_drawer_open, R.string.navigation_drawer_close);
        drawer.setDrawerListener(toggle);
        toggle.syncState();

        NavigationView navigationView = (NavigationView) findViewById(R.id.nav_view);
        navigationView.setNavigationItemSelectedListener(this);

        /* Functions for image gallery
        initImageLoader();
        init();
        */
    }
    private Uri getOutputMediaFileUri(int type){
        return Uri.fromFile(getOutputMediaFile(type));
    }

    private File getOutputMediaFile(int type){
        // To be safe, you should check that the SDCard is mounted
        // using Environment.getExternalStorageState() before doing this.
        EditText editStrata = (EditText) findViewById(R.id.editStrata);
        EditText editBuildingName = (EditText) findViewById(R.id.editBuildingName);
        EditText editFile = (EditText) findViewById(R.id.editFile);
        Spinner inspectionby = (Spinner) findViewById(R.id.spinnerInspectionBy);
        String Strata = editStrata.getText().toString();
        String Bname = editBuildingName.getText().toString();

        Calendar c = Calendar.getInstance();
        System.out.println("Current time => " + c.getTime());

        SimpleDateFormat df = new SimpleDateFormat("MMM-dd-yy");
        String formattedDate = df.format(c.getTime());
        File mediaStorageDir = new File(Environment.getExternalStoragePublicDirectory(
                Environment.DIRECTORY_PICTURES),  "Photos_F#" + editFile.getText().toString() + "_" + inspectionby.getSelectedItem().toString() + "_" + formattedDate.toString());
        // This location works best if you want the created images to be shared
        // between applications and persist after your app has been uninstalled.

        // Create the storage directory if it does not exist
        if (! mediaStorageDir.exists()){
            if (! mediaStorageDir.mkdirs()){
                Log.d("MyCameraApp", "failed to create directory");
                return null;
            }
        }

        // Create a media file name
        String timeStamp = new SimpleDateFormat("MMM-dd-yy-HH:mm.ss").format(new Date());
        File mediaFile;
         if (type == MEDIA_TYPE_IMAGE && Bname.matches("")){
            mediaFile = new File(mediaStorageDir.getPath() + File.separator +
                    editStrata.getText().toString() +  "_F#" + editFile.getText().toString() + "_" + inspectionby.getSelectedItem().toString()+ timeStamp + ".jpg");


        } else if (type == MEDIA_TYPE_IMAGE && Strata.matches("")) {
             mediaFile = new File(mediaStorageDir.getPath() + File.separator +
                     editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() + "_" + inspectionby.getSelectedItem().toString() + timeStamp + ".jpg");

         }else if (type == MEDIA_TYPE_IMAGE && Strata.matches("") && Bname.matches("")) {
             mediaFile = new File(mediaStorageDir.getPath() + File.separator +
                      "F#" + editFile.getText().toString() + "_" + inspectionby.getSelectedItem().toString() + timeStamp + ".jpg");
         }

         else if (type == MEDIA_TYPE_IMAGE && editBuildingName != null && editStrata != null){
                 mediaFile = new File(mediaStorageDir.getPath() + File.separator +
                         editStrata.getText().toString() + "_" + editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() + "_" + inspectionby.getSelectedItem().toString()+ timeStamp + ".jpg");

             } else{return null;}

        return mediaFile;
    }

    public void restartActivity(Activity act){
        Intent intent = new android.content.Intent();
        intent.setClass(act, MainActivity.class);
        act.startActivity(intent);
    }

    @Override
    public void onBackPressed() {
        DrawerLayout drawer = (DrawerLayout) findViewById(R.id.drawer_layout);
        if (drawer.isDrawerOpen(GravityCompat.START)) {
            drawer.closeDrawer(GravityCompat.START);
        } else {
            super.onBackPressed();
        }
    }

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        // Inflate the menu; this adds items to the action bar if it is present.
        getMenuInflater().inflate(R.menu.main, menu);
        return true;
    }

    @SuppressWarnings("StatementWithEmptyBody")
    @Override
    public boolean onNavigationItemSelected(MenuItem item) {
        // Handle navigation view item clicks here.
        int id = item.getItemId();

        if (id == R.id.nav_save) {
            saveExcelFile(this,"report.xls");

        }

        if (id == R.id.nav_share) {
            Intent i=
                    new Intent(this, MainActivity.class)
                            .setFlags(
                                    Intent.FLAG_ACTIVITY_NEW_TASK |

                                            Intent.FLAG_ACTIVITY_MULTIPLE_TASK);
            startActivity(i);

        }
        if (id == R.id.nav_clear) {
            new AlertDialog.Builder(MainActivity.this)
                    .setTitle("Clear report")
                    .setMessage("Are you sure you want to clear the report?")
                    .setPositiveButton(android.R.string.yes, new DialogInterface.OnClickListener() {
                        public void onClick(DialogInterface dialog, int which) {
                            restartActivity(MainActivity.this);
                        }
                    })
                    .setNegativeButton(android.R.string.no, new DialogInterface.OnClickListener() {
                        public void onClick(DialogInterface dialog, int which) {
                            dialog.dismiss();
                        }
                    })
                    .show();
        }

        DrawerLayout drawer = (DrawerLayout) findViewById(R.id.drawer_layout);
        drawer.closeDrawer(GravityCompat.START);
        return true;
    }

    private void saveExcelFile(Context context, String filename) {

        //GENERAL REPORT HIDDEN EDIT TEXTS
        EditText editCity = (EditText) findViewById(R.id.editCity);
        EditText editProvince = (EditText) findViewById(R.id.editProvince);
        EditText editAppraisalType = (EditText) findViewById(R.id.editAppraisalType);
        EditText editReportBy = (EditText) findViewById(R.id.editReportBy);
        EditText editInspectionBy = (EditText) findViewById(R.id.editInspectionBy);
        EditText editAssistedBy = (EditText) findViewById(R.id.editAssistedBy);
        //GENEREAL REPORT INFO  EDIT TEXTS
        EditText editFile = (EditText) findViewById(R.id.editFile);
        EditText editStrata = (EditText) findViewById(R.id.editStrata);
        EditText editBuildingName = (EditText) findViewById(R.id.editBuildingName);
        EditText editStAdd = (EditText) findViewById(R.id.editStAdd);
        EditText editInspectionDate = (EditText) findViewById(R.id.editInspectionDate);
        EditText editEffectiveDate = (EditText) findViewById(R.id.editEffectiveDate);
        EditText editCIV = (EditText) findViewById(R.id.editCIV);
        EditText editMIV = (EditText) findViewById(R.id.editMV);
        EditText editSiteContactName = (EditText) findViewById(R.id.editSiteContactName);
        EditText editSiteContactSB = (EditText) findViewById(R.id.editSiteContactSB);
        EditText editSiteContactPhone = (EditText) findViewById(R.id.editSiteContactPhone);
        EditText editResiSuitesNo   = (EditText) findViewById(R.id.editResiSuitesNo);
        EditText editCommUnitsNo   = (EditText) findViewById(R.id.editCommUnitsNo );
        EditText editBuildingNo   = (EditText) findViewById(R.id.editBuildingNo);
        EditText editConstYear  = (EditText) findViewById(R.id.editConstYear );
        EditText editZoning  = (EditText) findViewById(R.id.editZoning );
        EditText editParkSur   = (EditText) findViewById(R.id.editParkSurf);
        EditText editParkUG   = (EditText) findViewById(R.id.editParkUG );
        EditText editParkGRG  = (EditText) findViewById(R.id.editParkGRG);
        EditText editParkCP  = (EditText) findViewById(R.id.editParkCP);
        String Strata = editStrata.getText().toString();
        String Bname = editBuildingName.getText().toString();
        //SITE FEATURES EDIT TEXTS
        EditText editSiteArea  = (EditText) findViewById(R.id.editSiteArea );
        EditText editSoftPerc  = (EditText) findViewById(R.id.editSoftPerc );
        EditText editHardPerc  = (EditText) findViewById(R.id.editHardPerc );
        //COMMON BUILDING APPLIANCES EDIT TEXTS
        EditText editIntCmnApp_chkbx  = (EditText) findViewById(R.id.editIntCmnApp_chkbx );
        String garbcompuni = editIntCmnApp_chkbx.getText().toString();
        EditText editIntCmnApp_chkbx1  = (EditText) findViewById(R.id.editIntCmnApp_chkbx1 );
        String commwashuni = editIntCmnApp_chkbx1.getText().toString();
        EditText editIntCmnApp_chkbx2   = (EditText) findViewById(R.id.editIntCmnApp_chkbx2);
        String commdryuni = editIntCmnApp_chkbx2.getText().toString();

        EditText editMechElev_chkbx8  = (EditText) findViewById(R.id.editMechElev_chkbx8);
        EditText editMechElev_chkbx9  = (EditText) findViewById(R.id.editMechElev_chkbx9);

        //SECURITY EDIT TEXTS
        EditText editSecSysCamNo_chkbx   = (EditText) findViewById(R.id.editSecSysCamNo_chkbx );
        EditText editSecSysParkadeNo_chkbx  = (EditText) findViewById(R.id.editSecSysParkadeNo_chkbx);
        EditText editSecSysCamBk_chkbx  = (EditText) findViewById(R.id.editSecSysCamBk_chkbx);
        //FIREPLACE QUANTITY VALUE and dropdown
        EditText fireplacequan  = (EditText) findViewById(R.id.editIntApp_chkbx1);
        Spinner fireplacespinner = (Spinner) findViewById(R.id.spinnerIntIndvFireplace);





        //SITE FEATURES EDIT TEXTS CHECKBOXES
        EditText editlndscp_chkbx30  = (EditText) findViewById(R.id.editlndscp_chkbx30 );
        String box1L = editlndscp_chkbx30.getText().toString();
        EditText editLndscp_chkbx2 = (EditText) findViewById(R.id.editLndscp_chkbx2 );
        String box2L = editLndscp_chkbx2.getText().toString();
        EditText editLndscp_chkbx3  = (EditText) findViewById(R.id.editLndscp_chkbx3 );
        String box3L = editLndscp_chkbx3.getText().toString();
        EditText editLndscp_chkbx4  = (EditText) findViewById(R.id.editLndscp_chkbx4);
        String box4L = editLndscp_chkbx4.getText().toString();
        EditText editLndscp_chkbx5  = (EditText) findViewById(R.id.editLndscp_chkbx5 );
        String box5L = editLndscp_chkbx5.getText().toString();
        EditText editLndscp_chkbx6  = (EditText) findViewById(R.id.editLndscp_chkbx6 );
        String box6L = editLndscp_chkbx6.getText().toString();
        EditText editLndscp_chkbx7 = (EditText) findViewById(R.id.editLndscp_chkbx7 );
        String box7L = editLndscp_chkbx7.getText().toString();
        EditText editLndscp_chkbx8  = (EditText) findViewById(R.id.editLndscp_chkbx8 );
        String box8L = editLndscp_chkbx8.getText().toString();
        EditText editLndscp_chkbx9  = (EditText) findViewById(R.id.editLndscp_chkbx9 );
        String box9L = editLndscp_chkbx9.getText().toString();/*
        EditText editLndscp_chkbx10  = (EditText) findViewById(R.id.editLndscp_chkbx10 );
        String box10L = editLndscp_chkbx10.getText().toString();*/








        //GENERAL FEATURES
        EditText numberoffloorsgen  = (EditText) findViewById(R.id.editFloorNo);

        EditText editFeat_chkbx1  = (EditText) findViewById(R.id.editFeat_chkbx1 );
        String box1PR = editFeat_chkbx1.getText().toString();
        EditText editFeat_chkbx2 = (EditText) findViewById(R.id.editFeat_chkbx2 );
        String box2PR = editFeat_chkbx2.getText().toString();

        //HVAC EDITCHECKBOX
        EditText editMechHvac_chkbx1  = (EditText) findViewById(R.id.editMechHvac_chkbx1 );
        String box1HVAC = editMechHvac_chkbx1.getText().toString();
        EditText editMechHvac_chkbx2 = (EditText) findViewById(R.id.editMechHvac_chkbx2 );
        String box2HVAC = editMechHvac_chkbx2.getText().toString();
        EditText editMechHvac_chkbx3 = (EditText) findViewById(R.id.editMechHvac_chkbx3 );
        String box3HVAC = editMechHvac_chkbx3.getText().toString();


        //SECURITYSYS EDITCHECKBOX
        EditText editSecSysNote_chkbx1  = (EditText) findViewById(R.id.editSecSysNote_chkbx1 );
        String box1SEC = editSecSysNote_chkbx1.getText().toString();
        EditText editSecSysNote_chkbx2 = (EditText) findViewById(R.id.editSecSysNote_chkbx2 );
        String box2SEC = editSecSysNote_chkbx2.getText().toString();
        EditText editSecSysNote_chkbx3 = (EditText) findViewById(R.id.editSecSysNote_chkbx3 );
        String box3SEC = editSecSysNote_chkbx3.getText().toString();
        //FIRE EDITCHECKBOX
        EditText editFirePro_chkbx  = (EditText) findViewById(R.id.editFirePro_chkbx );
        String box1FIR = editFirePro_chkbx.getText().toString();
        EditText editFirePro_chkbx1 = (EditText) findViewById(R.id.editFirePro_chkbx1 );
        String box2FIR = editFirePro_chkbx1.getText().toString();
        EditText editFirePro_chkbx2 = (EditText) findViewById(R.id.editFirePro_chkbx2 );
        String box3FIR = editFirePro_chkbx2.getText().toString();




        //INDIVIDUAL SUITE APPLIANCES NOTES
        EditText editIntIndvNote  = (EditText) findViewById(R.id.editIntIndvNote );
        EditText editIntIndvNote1  = (EditText) findViewById(R.id.editIntIndvNote1 );
        //COMMON BUILDING APPLIANCES NOTES
        EditText editIntCmnAppNote  = (EditText) findViewById(R.id.editIntCmnAppNote );
        EditText editIntCmnAppNote1  = (EditText) findViewById(R.id.editIntCmnAppNote1 );
        //HVAC SYSTEM NOTES
        EditText editMechHvacNote  = (EditText) findViewById(R.id.editMechHvacNote );
        EditText editMechHvacNote2  = (EditText) findViewById(R.id.editMechHvacNote2 );
        //HVAC SYSTEM NOTES
        EditText editSecSysNote  = (EditText) findViewById(R.id.editSecSysNote );
        EditText editSecSysNote1  = (EditText) findViewById(R.id.editSecSysNote1 );
        //FIRE PROTECTION NOTES
        EditText editFireProNote  = (EditText) findViewById(R.id.editFireProNote );
        EditText editFireProNote1  = (EditText) findViewById(R.id.editFireProNote1 );
        //ELEC NOTES
        EditText editMechElec_chkbx1  = (EditText) findViewById(R.id.editMechElec_chkbx1 );
        //MAIN NOTES
        EditText editNotes  = (EditText) findViewById(R.id.editNotes );
        EditText editNotes1  = (EditText) findViewById(R.id.editNotes1 );
        EditText editNotes2  = (EditText) findViewById(R.id.editNotes2 );
        EditText editNotes3  = (EditText) findViewById(R.id.editNotes3 );



        //CHECKBOXES
        //SITE FEATURES
        CheckBox grass = (CheckBox) findViewById(R.id.lndscp_chkbx1);
        CheckBox surfacingparking = (CheckBox) findViewById(R.id.lndscp_chkbx27);
        CheckBox lighting = (CheckBox) findViewById(R.id.lndscp_chkbx14);
        CheckBox tenniscourt = (CheckBox) findViewById(R.id.lndscp_chkbx28);
        CheckBox trees = (CheckBox) findViewById(R.id.lndscp_chkbx2);
        CheckBox mailkiosk = (CheckBox) findViewById(R.id.lndscp_chkbx15);
        CheckBox walkways = (CheckBox) findViewById(R.id.lndscp_chkbx29);
        CheckBox scrubs = (CheckBox) findViewById(R.id.lndscp_chkbx3);
        CheckBox metalpoles = (CheckBox) findViewById(R.id.lndscp_chkbx16);
        CheckBox waterfeature = (CheckBox) findViewById(R.id.lndscp_chkbx30);
        CheckBox irrigationsys = (CheckBox) findViewById(R.id.lndscp_chkbx4);
        CheckBox outdoorspool  = (CheckBox) findViewById(R.id.lndscp_chkbx17);
        CheckBox bballcourt = (CheckBox) findViewById(R.id.lndscp_chkbx6);
        CheckBox patios = (CheckBox) findViewById(R.id.lndscp_chkbx18);
        CheckBox benches = (CheckBox) findViewById(R.id.lndscp_chkbx5);
        CheckBox planters = (CheckBox) findViewById(R.id.lndscp_chkbx19);
        CheckBox bikerack = (CheckBox) findViewById(R.id.lndscp_chkbx7);
        CheckBox playground = (CheckBox) findViewById(R.id.lndscp_chkbx20);
        CheckBox curbing = (CheckBox) findViewById(R.id.lndscp_chkbx8);
        CheckBox pedestriangates = (CheckBox) findViewById(R.id.lndscp_chkbx21);
        CheckBox entrancecanopy = (CheckBox) findViewById(R.id.lndscp_chkbx9);
        CheckBox railing = (CheckBox) findViewById(R.id.lndscp_chkbx22);
        CheckBox entranceg = (CheckBox) findViewById(R.id.lndscp_chkbx10);
        CheckBox retainingwalls = (CheckBox) findViewById(R.id.lndscp_chkbx23);
        CheckBox fencing = (CheckBox) findViewById(R.id.lndscp_chkbx11);
        CheckBox roadways = (CheckBox) findViewById(R.id.lndscp_chkbx24);
        CheckBox garbageshelter= (CheckBox) findViewById(R.id.lndscp_chkbx12);
        CheckBox signs = (CheckBox) findViewById(R.id.lndscp_chkbx25);
        CheckBox gazebo = (CheckBox) findViewById(R.id.lndscp_chkbx13);
        CheckBox storageshed = (CheckBox) findViewById(R.id.lndscp_chkbx26);
        //FEATURES
        CheckBox amenitybuilding = (CheckBox) findViewById(R.id.feat_chkbx1);
        CheckBox gym = (CheckBox) findViewById(R.id.feat_chkbx12);
        CheckBox partyroom = (CheckBox) findViewById(R.id.feat_chkbx23);
        CheckBox amenityroom = (CheckBox) findViewById(R.id.feat_chkbx2);
        CheckBox indoorhott = (CheckBox) findViewById(R.id.feat_chkbx13);
        CheckBox readingroom = (CheckBox) findViewById(R.id.feat_chkbx24);
        CheckBox bikeroom = (CheckBox) findViewById(R.id.feat_chkbx3);
        CheckBox indoorswimp = (CheckBox) findViewById(R.id.feat_chkbx14);
        CheckBox roofg = (CheckBox) findViewById(R.id.feat_chkbx25);
        CheckBox carpports = (CheckBox) findViewById(R.id.feat_chkbx4);
        CheckBox laundryroom  = (CheckBox) findViewById(R.id.feat_chkbx15);
        CheckBox sauna = (CheckBox) findViewById(R.id.feat_chkbx26);
        CheckBox changer = (CheckBox) findViewById(R.id.feat_chkbx5);
        CheckBox library = (CheckBox) findViewById(R.id.feat_chkbx16);
        CheckBox steamroom = (CheckBox) findViewById(R.id.feat_chkbx27);
        CheckBox commonkitch = (CheckBox) findViewById(R.id.feat_chkbx6);
        CheckBox mechroom = (CheckBox) findViewById(R.id.feat_chkbx17);
        CheckBox storage = (CheckBox) findViewById(R.id.feat_chkbx28);
        CheckBox commonwash = (CheckBox) findViewById(R.id.feat_chkbx7);
        CheckBox meetingr = (CheckBox) findViewById(R.id.feat_chkbx18);
        CheckBox workshop = (CheckBox) findViewById(R.id.feat_chkbx29);
        CheckBox crawlspace = (CheckBox) findViewById(R.id.feat_chkbx8);
        CheckBox multipur = (CheckBox) findViewById(R.id.feat_chkbx19);
        CheckBox andvar = (CheckBox) findViewById(R.id.feat_chkbx30);
        CheckBox exerciser = (CheckBox) findViewById(R.id.feat_chkbx9);
        CheckBox office= (CheckBox) findViewById(R.id.feat_chkbx20);
        CheckBox garagedetached = (CheckBox) findViewById(R.id.feat_chkbx10);
        CheckBox parkadeconc = (CheckBox) findViewById(R.id.feat_chkbx21);
        CheckBox garageatt = (CheckBox) findViewById(R.id.feat_chkbx11);
        CheckBox parkadeground = (CheckBox) findViewById(R.id.feat_chkbx22);
        //EXTRAS
        CheckBox balconies = (CheckBox) findViewById(R.id.feat_chkbx31);
        CheckBox roofdeck= (CheckBox) findViewById(R.id.feat_chkbx33);
        CheckBox skylights = (CheckBox) findViewById(R.id.feat_chkbx35);
        CheckBox canopies = (CheckBox) findViewById(R.id.feat_chkbx32);
        CheckBox roofoverhang = (CheckBox) findViewById(R.id.feat_chkbx34);
        //INDIVIDUALSUITE APPLIANCES
        CheckBox dishwasher = (CheckBox) findViewById(R.id.intApp_chkbx1);
        CheckBox stoveelectr = (CheckBox) findViewById(R.id.intApp_chkbx7);
        CheckBox dryer = (CheckBox) findViewById(R.id.intApp_chkbx13);
        CheckBox fridge = (CheckBox) findViewById(R.id.intApp_chkbx2);
        CheckBox stovegas = (CheckBox) findViewById(R.id.intApp_chkbx8);
        CheckBox washer = (CheckBox) findViewById(R.id.intApp_chkbx14);
        CheckBox garburator = (CheckBox) findViewById(R.id.intApp_chkbx3);
        CheckBox stovetopelec = (CheckBox) findViewById(R.id.intApp_chkbx9);
        CheckBox microwavewhood = (CheckBox) findViewById(R.id.intApp_chkbx4);
        CheckBox wallovenelec = (CheckBox) findViewById(R.id.intApp_chkbx10);
        CheckBox microwavenohood  = (CheckBox) findViewById(R.id.intApp_chkbx5);
        CheckBox stovetopgas = (CheckBox) findViewById(R.id.intApp_chkbx11);
        CheckBox centralvac = (CheckBox) findViewById(R.id.intApp_chkbx6);
        CheckBox wallovengas = (CheckBox) findViewById(R.id.intApp_chkbx12);
        CheckBox fireplace = (CheckBox) findViewById(R.id.intApp_chkbx15);
        //COMMON BUILDING APPLIANCES
        CheckBox garbagecompactor = (CheckBox) findViewById(R.id.intCmnApp_chkbx);
        CheckBox commonwashers = (CheckBox) findViewById(R.id.intCmnApp_chkbx1);
        CheckBox commondryers = (CheckBox) findViewById(R.id.intCmnApp_chkbx2);
        //HVAC SYSTEMS
        CheckBox aircondiioninghvac = (CheckBox) findViewById(R.id.mechHvac_chkbx1);
        CheckBox heatpumphvac = (CheckBox) findViewById(R.id.mechHvac_chkbx6);
        CheckBox spaceheaterhvac = (CheckBox) findViewById(R.id.mechHvac_chkbx11);
        CheckBox ceilingfanshvac = (CheckBox) findViewById(R.id.mechHvac_chkbx2);
        CheckBox hotwaterbaseboard = (CheckBox) findViewById(R.id.mechHvac_chkbx7);
        CheckBox ventsys = (CheckBox) findViewById(R.id.mechHvac_chkbx12);
        CheckBox elecbaseboard = (CheckBox) findViewById(R.id.mechHvac_chkbx3);
        CheckBox hotwaterradiant = (CheckBox) findViewById(R.id.mechHvac_chkbx8);
        CheckBox forcedair = (CheckBox) findViewById(R.id.mechHvac_chkbx4);
        CheckBox hotwaterradiator = (CheckBox) findViewById(R.id.mechHvac_chkbx9);
        CheckBox geothermal  = (CheckBox) findViewById(R.id.mechHvac_chkbx5);
        CheckBox rooftophvacunits = (CheckBox) findViewById(R.id.mechHvac_chkbx10);
        //Electrical system
        CheckBox circuitbreakers = (CheckBox) findViewById(R.id.mechElec_chkbx1);
        CheckBox fuseboxes  = (CheckBox) findViewById(R.id.mechElec_chkbx2);
        CheckBox elecEstim = (CheckBox) findViewById(R.id.mechElec_chkbx3);
        //Plumbing System
        CheckBox copper = (CheckBox) findViewById(R.id.mechElec_chkbx4);
        CheckBox PEX  = (CheckBox) findViewById(R.id.mechElec_chkbx5);
        CheckBox plumest = (CheckBox) findViewById(R.id.mechElec_chkbx6);
        //FIRE PROTECTION
        CheckBox smokedetwired = (CheckBox) findViewById(R.id.firePro_chkbx1);
        CheckBox smokedetbatt = (CheckBox) findViewById(R.id.firePro_chkbx9);
        CheckBox heatdetwired = (CheckBox) findViewById(R.id.firePro_chkbx2);
        CheckBox heatdetbatt = (CheckBox) findViewById(R.id.firePro_chkbx10);
        CheckBox firepanel = (CheckBox) findViewById(R.id.firePro_chkbx3);
        CheckBox firealarms = (CheckBox) findViewById(R.id.firePro_chkbx4);
        CheckBox handicapaccess = (CheckBox) findViewById(R.id.firePro_chkbx11);
        CheckBox pullstations = (CheckBox) findViewById(R.id.firePro_chkbx5);
        CheckBox elevramp = (CheckBox) findViewById(R.id.firePro_chkbx12);
        CheckBox fireexting = (CheckBox) findViewById(R.id.firePro_chkbx6);
        CheckBox emerglights  = (CheckBox) findViewById(R.id.firePro_chkbx7);
        CheckBox generator = (CheckBox) findViewById(R.id.firePro_chkbx8);
        //SPRINKLERS SYSTEM
        RadioButton nosprink = (RadioButton) findViewById(R.id.sprSys_radiobtn);
        RadioButton sprinklersthroughout = (RadioButton) findViewById(R.id.sprSys_radiobtn1);
        RadioButton parkadeboileronly  = (RadioButton) findViewById(R.id.sprSys_radiobtn2);
        RadioButton firewallsTHs = (RadioButton) findViewById(R.id.sprSys_radiobtn3);
        //STANDPIPE SYSTEM
        CheckBox nostandpipe = (CheckBox) findViewById(R.id.standSys_chkbx);
        CheckBox stairwells  = (CheckBox) findViewById(R.id.standSys_chkbx1);
        CheckBox wallcabinets = (CheckBox) findViewById(R.id.standSys_chkbx2);
        //SECURITY SYSTEMS
        CheckBox prewiredalarm = (CheckBox) findViewById(R.id.secSys_chkbx1);
        CheckBox enterphonesystem = (CheckBox) findViewById(R.id.secSys_chkbx4);
        CheckBox intercomsystem = (CheckBox) findViewById(R.id.secSys_chkbx6);
        CheckBox keyscanaccess = (CheckBox) findViewById(R.id.secSys_chkbx2);
        CheckBox videosurv = (CheckBox) findViewById(R.id.secSys_chkbx5);
        CheckBox parkadegates = (CheckBox) findViewById(R.id.secSys_chkbx3);



        //GENERAL REPORT INFORMATION SPPINERS
        Spinner spinnercity = (Spinner) findViewById(R.id.spinnerCity);
        String city = spinnercity.getSelectedItem().toString();
        Spinner spinnerprov = (Spinner) findViewById(R.id.spinnerProvince);
        String province = spinnerprov.getSelectedItem().toString();
        Spinner spinnerapptype = (Spinner) findViewById(R.id.spinnerAppraisalType);
        String apptype = spinnerapptype.getSelectedItem().toString();
        Spinner spinnerreportby = (Spinner) findViewById(R.id.spinnerReportBy);
        String reportby = spinnerreportby.getSelectedItem().toString();
        Spinner inspectionby = (Spinner) findViewById(R.id.spinnerInspectionBy);
        String inpectionbyT = inspectionby.getSelectedItem().toString();
        Spinner assistedby = (Spinner) findViewById(R.id.spinnerAssistedBy);
        String assisBy = assistedby.getSelectedItem().toString();
        //LANDSCAPING + GENERAL SPINNERS Structures
        Spinner softqua = (Spinner) findViewById(R.id.spinnerSoftQua);
        Spinner hardqua = (Spinner) findViewById(R.id.spinnerHardQua);
        Spinner buildingframe = (Spinner) findViewById(R.id.spinnerBuildingFrame);
        Spinner buildingtypeuse = (Spinner) findViewById(R.id.spinnerBuildingType1);
        Spinner buildingtypeperce = (Spinner) findViewById(R.id.spinnerBuildingType2);
        Spinner propertytype = (Spinner) findViewById(R.id.spinnerPropType);
        Spinner buildingclass = (Spinner) findViewById(R.id.spinnerBuildClass);
        Spinner bcbcgroup = (Spinner) findViewById(R.id.spinnerBcbcGroup);
        Spinner floordropdown = (Spinner) findViewById(R.id.spinnerFeatFloorNo);
        Spinner floordropdownother = (Spinner) findViewById(R.id.spinnerFeatFloorNoOther);
        Spinner featproprights = (Spinner) findViewById(R.id.spinnerFeatPropRights);
        //INDIVUAL SUITEAPPLIANCES SPINNERS
        Spinner spinnerIntAppBrand = (Spinner) findViewById(R.id.spinnerIntAppBrand);
        Spinner spinnerIntAppBrand2 = (Spinner) findViewById(R.id.spinnerIntAppBrand2);
        Spinner spinnerIntAppQuality = (Spinner) findViewById(R.id.spinnerIntAppQuality);
        Spinner spinnerIntAppQuality2 = (Spinner) findViewById(R.id.spinnerIntAppQuality2);
        Spinner spinnerIntAppStyle = (Spinner) findViewById(R.id.spinnerIntAppStyle);
        Spinner spinnerIntAppStyle2 = (Spinner) findViewById(R.id.spinnerIntAppStyle2);
        //COMMONBUILDINGAPPLIANCES SPINNERS

        Spinner spinnerCmnBuildApp = (Spinner) findViewById(R.id.spinnerCmnBuildApp);
        String buildap = spinnerCmnBuildApp.getSelectedItem().toString();
        Spinner spinnerCmnBuildApp1 = (Spinner) findViewById(R.id.spinnerCmnBuildApp1);
        String buildap1 = spinnerCmnBuildApp1.getSelectedItem().toString();
        Spinner spinnerCmnBuildApp2 = (Spinner) findViewById(R.id.spinnerCmnBuildApp2);
        String buildap2 = spinnerCmnBuildApp2.getSelectedItem().toString();
        Spinner spinnerMechElev = (Spinner) findViewById(R.id.spinnerMechElev);
        String Smechelev = spinnerMechElev.getSelectedItem().toString();
        //Finnish
        //INTERIORFINISHES: CEILING
        Spinner spinnerIntCeilingPark = (Spinner) findViewById(R.id.spinnerIntCeilingPark);
        String cPark = spinnerIntCeilingPark.getSelectedItem().toString();
        Spinner spinnerIntCeilingPark1 = (Spinner) findViewById(R.id.spinnerIntCeilingPark1);
        String cPark1 = spinnerIntCeilingPark1.getSelectedItem().toString();
        Spinner spinnerIntCeilingPark2 = (Spinner) findViewById(R.id.spinnerIntCeilingPark2);
        String cPark2 = spinnerIntCeilingPark2.getSelectedItem().toString();
        Spinner spinnerIntCeilingComm = (Spinner) findViewById(R.id.spinnerIntCeilingComm);
        String cCom = spinnerIntCeilingComm.getSelectedItem().toString();
        Spinner spinnerIntCeilingComm1 = (Spinner) findViewById(R.id.spinnerIntCeilingComm1);
        String cCom1 = spinnerIntCeilingComm1.getSelectedItem().toString();
        Spinner spinnerIntCeilingComm2 = (Spinner) findViewById(R.id.spinnerIntCeilingComm2);
        String cCom2 = spinnerIntCeilingComm2.getSelectedItem().toString();
        Spinner spinnerIntCeilingBase = (Spinner) findViewById(R.id.spinnerIntCeilingBase);
        String cBas = spinnerIntCeilingBase.getSelectedItem().toString();
        Spinner spinnerIntCeilingBase1 = (Spinner) findViewById(R.id.spinnerIntCeilingBase1);
        String cBas1 = spinnerIntCeilingBase1.getSelectedItem().toString();
        Spinner spinnerIntCeilingBase2 = (Spinner) findViewById(R.id.spinnerIntCeilingBase2);
        String cBas2 = spinnerIntCeilingBase2.getSelectedItem().toString();
        Spinner spinnerIntCeilingLobby = (Spinner) findViewById(R.id.spinnerIntCeilingLobby);
        String cLob = spinnerIntCeilingLobby.getSelectedItem().toString();
        Spinner spinnerIntCeilingLobby1 = (Spinner) findViewById(R.id.spinnerIntCeilingLobby1);
        String cLob1 = spinnerIntCeilingLobby1.getSelectedItem().toString();
        Spinner spinnerIntCeilingLobby2 = (Spinner) findViewById(R.id.spinnerIntCeilingLobby2);
        String cLob2 = spinnerIntCeilingLobby2.getSelectedItem().toString();
        Spinner spinnerIntCeilingHall = (Spinner) findViewById(R.id.spinnerIntCeilingHall);
        String cHal = spinnerIntCeilingHall.getSelectedItem().toString();
        Spinner spinnerIntCeilingHall1 = (Spinner) findViewById(R.id.spinnerIntCeilingHall1);
        String cHal1 = spinnerIntCeilingHall1.getSelectedItem().toString();
        Spinner spinnerIntCeilingHall2 = (Spinner) findViewById(R.id.spinnerIntCeilingHall2);
        String cHa2 = spinnerIntCeilingHall2.getSelectedItem().toString();
        Spinner spinnerIntCeilingSuite = (Spinner) findViewById(R.id.spinnerIntCeilingSuite);
        String cSui = spinnerIntCeilingSuite.getSelectedItem().toString();
        Spinner spinnerIntCeilingSuite1 = (Spinner) findViewById(R.id.spinnerIntCeilingSuite1);
        String cSui1 = spinnerIntCeilingSuite1.getSelectedItem().toString();
        Spinner spinnerIntCeilingSuite2 = (Spinner) findViewById(R.id.spinnerIntCeilingSuite2);
        String cSui2 = spinnerIntCeilingSuite2.getSelectedItem().toString();
        Spinner spinnerIntCeilingOther = (Spinner) findViewById(R.id.spinnerIntCeilingOther);
        String cOth = spinnerIntCeilingOther.getSelectedItem().toString();
        Spinner spinnerIntCeilingOther1 = (Spinner) findViewById(R.id.spinnerIntCeilingOther1);
        String cOth1 = spinnerIntCeilingOther1.getSelectedItem().toString();
        Spinner spinnerIntCeilingOther2 = (Spinner) findViewById(R.id.spinnerIntCeilingOther2);
        String cOth2 = spinnerIntCeilingOther2.getSelectedItem().toString();
        Spinner spinnerIntCeilingOther3 = (Spinner) findViewById(R.id.spinnerIntCeilingOther3);
        String cOth3 = spinnerIntCeilingOther3.getSelectedItem().toString();
        Spinner spinnerIntCeilingOther4 = (Spinner) findViewById(R.id.spinnerIntCeilingOther4);
        String cOth4 = spinnerIntCeilingOther4.getSelectedItem().toString();
        Spinner spinnerIntCeilingOther5 = (Spinner) findViewById(R.id.spinnerIntCeilingOther5);
        String cOth5 = spinnerIntCeilingOther5.getSelectedItem().toString();
        //INTERIORFINISHES: WALLS
        Spinner spinnerIntWallsPark = (Spinner) findViewById(R.id.spinnerIntWallsPark);
        String wPar = spinnerIntWallsPark.getSelectedItem().toString();
        Spinner spinnerIntWallsPark1 = (Spinner) findViewById(R.id.spinnerIntWallsPark1);
        String wPar1 = spinnerIntWallsPark1.getSelectedItem().toString();
        Spinner spinnerIntWallsPark2 = (Spinner) findViewById(R.id.spinnerIntWallsPark2);
        String wPar2 = spinnerIntWallsPark2.getSelectedItem().toString();
        Spinner spinnerIntWallsComm = (Spinner) findViewById(R.id.spinnerIntWallsComm);
        String wCom= spinnerIntWallsComm.getSelectedItem().toString();
        Spinner spinnerIntWallsComm1 = (Spinner) findViewById(R.id.spinnerIntWallsComm1);
        String wCom1 = spinnerIntWallsComm1.getSelectedItem().toString();
        Spinner spinnerIntWallsComm2 = (Spinner) findViewById(R.id.spinnerIntWallsComm2);
        String wCom2= spinnerIntWallsComm2.getSelectedItem().toString();
        Spinner spinnerIntWallsBase = (Spinner) findViewById(R.id.spinnerIntWallsBase);
        String wBas= spinnerIntWallsBase.getSelectedItem().toString();
        Spinner spinnerIntWallsBase1 = (Spinner) findViewById(R.id.spinnerIntWallsBase1);
        String wBas1= spinnerIntWallsBase1.getSelectedItem().toString();
        Spinner spinnerIntWallsBase2 = (Spinner) findViewById(R.id.spinnerIntWallsBase2);
        String wBas2= spinnerIntWallsBase2.getSelectedItem().toString();
        Spinner spinnerIntWallsLobby = (Spinner) findViewById(R.id.spinnerIntWallsLobby);
        String wLob = spinnerIntWallsLobby.getSelectedItem().toString();
        Spinner spinnerIntWallsLobby1 = (Spinner) findViewById(R.id.spinnerIntWallsLobby1);
        String wLob1 = spinnerIntWallsLobby1.getSelectedItem().toString();
        Spinner spinnerIntWallsLobby2 = (Spinner) findViewById(R.id.spinnerIntWallsLobby2);
        String wLob2 = spinnerIntWallsLobby2.getSelectedItem().toString();
        Spinner spinnerIntWallsHall = (Spinner) findViewById(R.id.spinnerIntWallsHall);
        String wHal = spinnerIntWallsHall.getSelectedItem().toString();
        Spinner spinnerIntWallsHall1 = (Spinner) findViewById(R.id.spinnerIntWallsHall1);
        String wHal1 = spinnerIntWallsHall1.getSelectedItem().toString();
        Spinner spinnerIntWallsHall2 = (Spinner) findViewById(R.id.spinnerIntWallsHall2);
        String wHal2 = spinnerIntWallsHall2.getSelectedItem().toString();
        Spinner spinnerIntWallsSuite = (Spinner) findViewById(R.id.spinnerIntWallsSuite);
        String wSui = spinnerIntWallsSuite.getSelectedItem().toString();
        Spinner spinnerIntWallsSuite1 = (Spinner) findViewById(R.id.spinnerIntWallsSuite1);
        String wSui1 = spinnerIntWallsSuite1.getSelectedItem().toString();
        Spinner spinnerIntWallsSuite2 = (Spinner) findViewById(R.id.spinnerIntWallsSuite2);
        String wSui2 = spinnerIntWallsSuite2.getSelectedItem().toString();
        Spinner spinnerIntWallsOther = (Spinner) findViewById(R.id.spinnerIntWallsOther);
        String wOth = spinnerIntWallsOther.getSelectedItem().toString();
        Spinner spinnerIntWallsOther1 = (Spinner) findViewById(R.id.spinnerIntWallsOther1);
        String wOth1 = spinnerIntWallsOther1.getSelectedItem().toString();
        Spinner spinnerIntWallsOther2 = (Spinner) findViewById(R.id.spinnerIntWallsOther2);
        String wOth2 = spinnerIntWallsOther2.getSelectedItem().toString();
        Spinner spinnerIntWallsOther3 = (Spinner) findViewById(R.id.spinnerIntWallsOther3);
        String wOth3 = spinnerIntWallsOther3.getSelectedItem().toString();
        Spinner spinnerIntWallsOther4 = (Spinner) findViewById(R.id.spinnerIntWallsOther4);
        String wOth4 = spinnerIntWallsOther4.getSelectedItem().toString();
        Spinner spinnerIntWallsOther5 = (Spinner) findViewById(R.id.spinnerIntWallsOther5);
        String wOth5 = spinnerIntWallsOther5.getSelectedItem().toString();
        //INTERIORFINISHES: FLOORS
        Spinner spinnerIntFloorPark = (Spinner) findViewById(R.id.spinnerIntFloorPark);
        String fPar = spinnerIntFloorPark.getSelectedItem().toString();
        Spinner spinnerIntFloorPark1 = (Spinner) findViewById(R.id.spinnerIntFloorPark1);
        String fPar1 = spinnerIntFloorPark1.getSelectedItem().toString();
        Spinner spinnerIntFloorPark2 = (Spinner) findViewById(R.id.spinnerIntFloorPark2);
        String fPar2 = spinnerIntFloorPark2.getSelectedItem().toString();
        Spinner spinnerIntFloorComm = (Spinner) findViewById(R.id.spinnerIntFloorComm);
        String fCom = spinnerIntFloorComm.getSelectedItem().toString();
        Spinner spinnerIntFloorComm1 = (Spinner) findViewById(R.id.spinnerIntFloorComm1);
        String fCom1 = spinnerIntFloorComm1.getSelectedItem().toString();
        Spinner spinnerIntFloorComm2 = (Spinner) findViewById(R.id.spinnerIntFloorComm2);
        String fCom2 = spinnerIntFloorComm2.getSelectedItem().toString();
        Spinner spinnerIntFloorBase = (Spinner) findViewById(R.id.spinnerIntFloorBase);
        String fBas = spinnerIntFloorBase.getSelectedItem().toString();
        Spinner spinnerIntFloorBase1 = (Spinner) findViewById(R.id.spinnerIntFloorBase1);
        String fBas1 = spinnerIntFloorBase1.getSelectedItem().toString();
        Spinner spinnerIntFloorBase2 = (Spinner) findViewById(R.id.spinnerIntFloorBase2);
        String fBas2 = spinnerIntFloorBase2.getSelectedItem().toString();
        Spinner spinnerIntFloorLobby = (Spinner) findViewById(R.id.spinnerIntFloorLobby);
        String fLob = spinnerIntFloorLobby.getSelectedItem().toString();
        Spinner spinnerIntFloorLobby1 = (Spinner) findViewById(R.id.spinnerIntFloorLobby1);
        String fLob1 = spinnerIntFloorLobby1.getSelectedItem().toString();
        Spinner spinnerIntFloorLobby2 = (Spinner) findViewById(R.id.spinnerIntFloorLobby2);
        String fLob2 = spinnerIntFloorLobby2.getSelectedItem().toString();
        Spinner spinnerIntFloorHall = (Spinner) findViewById(R.id.spinnerIntFloorHall);
        String fHal = spinnerIntFloorHall.getSelectedItem().toString();
        Spinner spinnerIntFloorHall1 = (Spinner) findViewById(R.id.spinnerIntFloorHall1);
        String fHal1 = spinnerIntFloorHall1.getSelectedItem().toString();
        Spinner spinnerIntFloorHall2 = (Spinner) findViewById(R.id.spinnerIntFloorHall2);
        String fHal2 = spinnerIntFloorHall2.getSelectedItem().toString();
        Spinner spinnerIntFloorSuite = (Spinner) findViewById(R.id.spinnerIntFloorSuite);
        String fSui = spinnerIntFloorSuite.getSelectedItem().toString();
        Spinner spinnerIntFloorSuite1 = (Spinner) findViewById(R.id.spinnerIntFloorSuite1);
        String fSui1 = spinnerIntFloorSuite1.getSelectedItem().toString();
        Spinner spinnerIntFloorSuite2 = (Spinner) findViewById(R.id.spinnerIntFloorSuite2);
        String fSui2 = spinnerIntFloorSuite2.getSelectedItem().toString();
        Spinner spinnerIntFloorOther = (Spinner) findViewById(R.id.spinnerIntFloorOther);
        String fOth = spinnerIntFloorOther.getSelectedItem().toString();
        Spinner spinnerIntFloorOther1 = (Spinner) findViewById(R.id.spinnerIntFloorOther1);
        String fOth1 = spinnerIntFloorOther1.getSelectedItem().toString();
        Spinner spinnerIntFloorOther2 = (Spinner) findViewById(R.id.spinnerIntFloorOther2);
        String fOth2 = spinnerIntFloorOther2.getSelectedItem().toString();
        Spinner spinnerIntFloorOther3 = (Spinner) findViewById(R.id.spinnerIntFloorOther3);
        String fOth3 = spinnerIntFloorOther3.getSelectedItem().toString();
        Spinner spinnerIntFloorOther4 = (Spinner) findViewById(R.id.spinnerIntFloorOther4);
        String fOth4 = spinnerIntFloorOther4.getSelectedItem().toString();
        Spinner spinnerIntFloorOther5 = (Spinner) findViewById(R.id.spinnerIntFloorOther5);
        String fOth5 = spinnerIntFloorOther5.getSelectedItem().toString();
        //FEATURES PRIMARY
        Spinner spinnerFeatExcav = (Spinner) findViewById(R.id.spinnerFeatExcav);
        String ExcaPri = spinnerFeatExcav.getSelectedItem().toString();
        Spinner spinnerFeatFound = (Spinner) findViewById(R.id.spinnerFeatFound);
        String FoundPrim = spinnerFeatFound.getSelectedItem().toString();
        Spinner spinnerExtFrame = (Spinner) findViewById(R.id.spinnerExtFrame);
        String ExtFramPrim = spinnerExtFrame.getSelectedItem().toString();
        Spinner spinnerIntFrame = (Spinner) findViewById(R.id.spinnerIntFrame);
        String IntFramPrim = spinnerIntFrame.getSelectedItem().toString();
        Spinner spinnerExtClad = (Spinner) findViewById(R.id.spinnerExtClad);
        String ExtCladPrim = spinnerExtClad.getSelectedItem().toString();
        Spinner spinnerLowFl = (Spinner) findViewById(R.id.spinnerLowFl);
        String lowFloorprim = spinnerLowFl.getSelectedItem().toString();
        Spinner spinnerUppFl = (Spinner) findViewById(R.id.spinnerUppFl);
        String uppFloorPrim = spinnerUppFl.getSelectedItem().toString();
        Spinner spinnerRoofStr = (Spinner) findViewById(R.id.spinnerRoofStr);
        String roofStrPrim = spinnerRoofStr.getSelectedItem().toString();
        Spinner spinnerRoofMat = (Spinner) findViewById(R.id.spinnerRoofMat);
        String roofMatPrim = spinnerRoofMat.getSelectedItem().toString();
        //FEATURES: ADDITIONAL 1
        Spinner spinnerFeatExcavAdd1 = (Spinner) findViewById(R.id.spinnerFeatExcavAdd1);
        String ExcaAdd1 = spinnerFeatExcavAdd1.getSelectedItem().toString();
        Spinner spinnerFeatFoundAdd1 = (Spinner) findViewById(R.id.spinnerFeatFoundAdd1);
        String FoundAdd1 = spinnerFeatFoundAdd1.getSelectedItem().toString();
        Spinner spinnerExtFrameAdd1 = (Spinner) findViewById(R.id.spinnerExtFrameAdd1);
        String ExtFramAdd1 = spinnerExtFrameAdd1.getSelectedItem().toString();
        Spinner spinnerIntFrameAdd1 = (Spinner) findViewById(R.id.spinnerIntFrameAdd1);
        String IntFramAdd1 = spinnerIntFrameAdd1.getSelectedItem().toString();
        Spinner spinnerExtCladAdd1 = (Spinner) findViewById(R.id.spinnerExtCladAdd1);
        String ExtCladAdd1 = spinnerExtCladAdd1.getSelectedItem().toString();
        Spinner spinnerLowFlAdd1 = (Spinner) findViewById(R.id.spinnerLowFlAdd1);
        String lowFloorAdd1 = spinnerLowFlAdd1.getSelectedItem().toString();
        Spinner spinnerUppFlAdd1 = (Spinner) findViewById(R.id.spinnerUppFlAdd1);
        String uppFloorAdd1 = spinnerUppFlAdd1.getSelectedItem().toString();
        Spinner spinnerRoofStrAdd1 = (Spinner) findViewById(R.id.spinnerRoofStrAdd1);
        String roofStrAdd1= spinnerRoofStrAdd1.getSelectedItem().toString();
        Spinner spinnerRoofMatAdd1 = (Spinner) findViewById(R.id.spinnerRoofMatAdd1);
        String roofMatAdd1 = spinnerRoofMatAdd1.getSelectedItem().toString();
        //FEATURES: ADDITIONAL 2
        Spinner spinnerFeatExcavAdd2 = (Spinner) findViewById(R.id.spinnerFeatExcavAdd2);
        String ExcaAdd2 = spinnerFeatExcavAdd2.getSelectedItem().toString();
        Spinner spinnerFeatFoundAdd2 = (Spinner) findViewById(R.id.spinnerFeatFoundAdd2);
        String FoundAdd2 = spinnerFeatFoundAdd2.getSelectedItem().toString();
        Spinner spinnerExtFrameAdd2 = (Spinner) findViewById(R.id.spinnerExtFrameAdd2);
        String ExtFramAdd2 = spinnerExtFrameAdd2.getSelectedItem().toString();
        Spinner spinnerIntFrameAdd2 = (Spinner) findViewById(R.id.spinnerIntFrameAdd2);
        String IntFramAdd2 = spinnerIntFrameAdd2.getSelectedItem().toString();
        Spinner spinnerExtCladAdd2 = (Spinner) findViewById(R.id.spinnerExtCladAdd2);
        String ExtCladAdd2 = spinnerExtCladAdd2.getSelectedItem().toString();
        Spinner spinnerLowFlAdd2 = (Spinner) findViewById(R.id.spinnerLowFlAdd2);
        String lowFloorAdd2 = spinnerLowFlAdd2.getSelectedItem().toString();
        Spinner spinnerUppFlAdd2 = (Spinner) findViewById(R.id.spinnerUppFlAdd2);
        String uppFloorAdd2 = spinnerUppFlAdd2.getSelectedItem().toString();
        Spinner spinnerRoofStrAdd2 = (Spinner) findViewById(R.id.spinnerRoofStrAdd2);
        String roofStrAdd2= spinnerRoofStrAdd2.getSelectedItem().toString();
        Spinner spinnerRoofMatAdd2 = (Spinner) findViewById(R.id.spinnerRoofMatAdd2);
        String roofMatAdd2 = spinnerRoofMatAdd2.getSelectedItem().toString();







        if (!isExternalStorageAvailable() || isExternalStorageReadOnly())
        {
            Log.e(TAG, "Storage not available or read only");
            return;
        }

        try{
            // Creating Input Stream
            FileInputStream myInput = new FileInputStream((new File(Environment.getExternalStoragePublicDirectory(
                    Environment.DIRECTORY_DOCUMENTS), filename)));

            // Create a POIFSFileSystem object
            // POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

            // Create a workbook using the File System
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myInput);

            // Get the first sheet from workbook
            HSSFSheet mySheet = myWorkBook.getSheetAt(0);
            Cell cell = null;








            /////////// GENERAL REPORT INFORMATIONSPINNERS/////////////////////////////////
            cell = mySheet.getRow(2).getCell(9);
            if (apptype.matches("Other")) {
                cell.setCellValue(editAppraisalType.getText().toString());
            } else {
                cell.setCellValue(spinnerapptype.getSelectedItem().toString());
            }

            cell = mySheet.getRow(4).getCell(9);
            if (reportby.matches("Other")) {
                cell.setCellValue(editReportBy.getText().toString());
            } else {
                cell.setCellValue(spinnerreportby.getSelectedItem().toString());
            }
            cell = mySheet.getRow(3).getCell(9);
            if (inpectionbyT.matches("Other")) {
                cell.setCellValue(editInspectionBy.getText().toString());
            } else {
                cell.setCellValue(inspectionby.getSelectedItem().toString());
            }

            cell = mySheet.getRow(5).getCell(9);
            if (assisBy.matches("Other")) {
                cell.setCellValue(editAssistedBy.getText().toString());
            } else {
                cell.setCellValue(assistedby.getSelectedItem().toString());
            }
            //LANDSCAPING + GENERAL SPINNERS Structures
            cell = mySheet.getRow(14).getCell(1);
            cell.setCellValue(softqua.getSelectedItem().toString());
            cell = mySheet.getRow(14).getCell(7);
            cell.setCellValue(hardqua.getSelectedItem().toString());
            cell = mySheet.getRow(27).getCell(3);
            cell.setCellValue(buildingframe.getSelectedItem().toString());
            cell = mySheet.getRow(28).getCell(3);
            cell.setCellValue(propertytype.getSelectedItem().toString());
            cell = mySheet.getRow(27).getCell(9);
            cell.setCellValue(buildingtypeuse.getSelectedItem().toString());
            cell = mySheet.getRow(28).getCell(9);
            cell.setCellValue(buildingtypeperce.getSelectedItem().toString());
            cell = mySheet.getRow(29).getCell(9);
            cell.setCellValue(bcbcgroup.getSelectedItem().toString());
            cell = mySheet.getRow(43).getCell(7);
            cell.setCellValue(buildingclass.getSelectedItem().toString());
            cell = mySheet.getRow(45).getCell(6);
            cell.setCellValue(floordropdown.getSelectedItem().toString());
            cell = mySheet.getRow(45).getCell(9);
            cell.setCellValue(floordropdownother.getSelectedItem().toString());
            cell = mySheet.getRow(46).getCell(3);
            cell.setCellValue(featproprights.getSelectedItem().toString());
            //INDIVUAL SUITEAPPLIANCES SPINNERS
            cell = mySheet.getRow(24).getCell(14);
            cell.setCellValue(spinnerIntAppBrand.getSelectedItem().toString());
            cell = mySheet.getRow(25).getCell(14);
            cell.setCellValue(spinnerIntAppBrand2.getSelectedItem().toString());
            cell = mySheet.getRow(24).getCell(18);
            cell.setCellValue(spinnerIntAppQuality.getSelectedItem().toString());
            cell = mySheet.getRow(25).getCell(18);
            cell.setCellValue(spinnerIntAppQuality2.getSelectedItem().toString());
            cell = mySheet.getRow(24).getCell(22);
            cell.setCellValue(spinnerIntAppStyle.getSelectedItem().toString());
            cell = mySheet.getRow(25).getCell(22);
            cell.setCellValue(spinnerIntAppStyle2.getSelectedItem().toString());
            //COMMONBUILDINGAPPLIANCES SPINNERS

            cell = mySheet.getRow(30).getCell(14);
            if (!buildap.matches("Select Option")) {
                cell.setCellValue(buildap);
            }

            cell = mySheet.getRow(30).getCell(18);
            if (!buildap1.matches("Select Option")) {
                cell.setCellValue(buildap1);
            }
            cell = mySheet.getRow(30).getCell(22);
            if (!buildap2.matches("Select Option")) {
                cell.setCellValue(buildap2);
            }
            cell = mySheet.getRow(34).getCell(14);
            if (!Smechelev.matches("Select Option")) {
                cell.setCellValue(spinnerMechElev.getSelectedItem().toString());
            }
            //FIREPLACE SPINNER QUANTITY
            cell = mySheet.getRow(21).getCell(22);
            cell.setCellValue(fireplacequan.getText().toString());
            cell = mySheet.getRow(21).getCell(23);
            cell.setCellValue(fireplacespinner.getSelectedItem().toString());

            //INTERIOR FINISHHHH
            //INTERIORFINISHES: CEILING

            cell = mySheet.getRow(4).getCell(14);
            if (!cPark.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingPark.getSelectedItem().toString());
            }

            cell = mySheet.getRow(5).getCell(14);
            if (!cPark1.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingPark1.getSelectedItem().toString());
            }

            cell = mySheet.getRow(6).getCell(14);
            if (!cPark2.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingPark2.getSelectedItem().toString());
            }

            cell = mySheet.getRow(4).getCell(15);
            if (!cCom.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingComm.getSelectedItem().toString());
            }

            cell = mySheet.getRow(5).getCell(15);
            if (!cCom1.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingComm1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(6).getCell(15);
            if (!cCom2.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingComm2.getSelectedItem().toString());
            }

            cell = mySheet.getRow(4).getCell(17);

            if (!cBas.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingBase.getSelectedItem().toString());
            }

            cell = mySheet.getRow(5).getCell(17);
            if (!cBas1.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingBase1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(6).getCell(17);
            if (!cBas2.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingBase2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(4).getCell(18);
            if (!cLob.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingLobby.getSelectedItem().toString());
            }

            cell = mySheet.getRow(5).getCell(18);
            if (!cLob1.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingLobby1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(6).getCell(18);
            if (!cLob2.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingLobby2.getSelectedItem().toString());
            }

            cell = mySheet.getRow(4).getCell(19);
            if (!cHal.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingHall.getSelectedItem().toString());
            }

            cell = mySheet.getRow(5).getCell(19);
            if (!cHal1.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingHall1.getSelectedItem().toString());
            }

            cell = mySheet.getRow(6).getCell(19);
            if (!cHa2.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingHall2.getSelectedItem().toString());
            }

            cell = mySheet.getRow(4).getCell(21);
            if (!cSui.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingSuite.getSelectedItem().toString());
            }

            cell = mySheet.getRow(5).getCell(21);
            if (!cSui1.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingSuite1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(6).getCell(21);
            if (!cSui2.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingSuite2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(4).getCell(22);
            if (!cOth.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingOther.getSelectedItem().toString());
            }

            cell = mySheet.getRow(5).getCell(22);
            if (!cOth1.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingOther1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(6).getCell(22);
            if (!cOth2.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingOther2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(4).getCell(23);
            if (!cOth3.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingOther3.getSelectedItem().toString());
            }
            cell = mySheet.getRow(5).getCell(23);
            if (!cOth4.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingOther4.getSelectedItem().toString());
            }
            cell = mySheet.getRow(6).getCell(23);

            if (!cOth5.matches("Select Option")) {
                cell.setCellValue(spinnerIntCeilingOther5.getSelectedItem().toString());
            }
            //INTERIORFINISHES: WALLS
            cell = mySheet.getRow(7).getCell(14);
            if (!wPar.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsPark.getSelectedItem().toString());
            }

            cell = mySheet.getRow(8).getCell(14);
            if (!wPar1.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsPark1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(9).getCell(14);
            if (!wPar2.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsPark2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(7).getCell(15);
            if (!wCom.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsComm.getSelectedItem().toString());
            }

            cell = mySheet.getRow(8).getCell(15);
            if (!wCom1.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsComm1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(9).getCell(15);
            if (!wCom2.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsComm2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(7).getCell(17);
            if (!wBas.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsBase.getSelectedItem().toString());
            }

            cell = mySheet.getRow(8).getCell(17);
            if (!wBas1.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsBase1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(9).getCell(17);
            if (!wBas2.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsBase2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(7).getCell(18);
            if (!wLob.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsLobby.getSelectedItem().toString());
            }
            cell = mySheet.getRow(8).getCell(18);
            if (!wLob1.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsLobby1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(9).getCell(18);
            if (!wLob2.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsLobby2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(7).getCell(19);

            if (!wHal.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsHall.getSelectedItem().toString());
            }

            cell = mySheet.getRow(8).getCell(19);
            if (!wHal1.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsHall1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(9).getCell(19);
            if (!wHal2.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsHall2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(7).getCell(21);
            if (!wSui.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsSuite.getSelectedItem().toString());
            }
            cell = mySheet.getRow(8).getCell(21);
            if (!wSui1.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsSuite1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(9).getCell(21);
            if (!wSui2.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsSuite2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(7).getCell(22);
            if (!wOth.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsOther.getSelectedItem().toString());
            }

            cell = mySheet.getRow(8).getCell(22);
            if (!wOth1.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsOther1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(9).getCell(22);
            if (!wOth2.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsOther2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(7).getCell(23);
            if (!wOth3.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsOther3.getSelectedItem().toString());
            }
            cell = mySheet.getRow(8).getCell(23);
            if (!wOth4.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsOther4.getSelectedItem().toString());
            }
            cell = mySheet.getRow(9).getCell(23);
            if (!wOth5.matches("Select Option")) {
                cell.setCellValue(spinnerIntWallsOther5.getSelectedItem().toString());
            }
            //INTERIORFINISHES: FLOORS
            cell = mySheet.getRow(10).getCell(14);
            if (!fPar.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorPark.getSelectedItem().toString());
            }

            cell = mySheet.getRow(11).getCell(14);
            if (!fPar1.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorPark1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(12).getCell(14);
            if (!fPar2.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorPark2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(10).getCell(15);
            if (!fCom.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorComm.getSelectedItem().toString());
            }
            cell = mySheet.getRow(11).getCell(15);
            if (!fCom1.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorComm1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(12).getCell(15);
            if (!fCom2.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorComm2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(10).getCell(17);
            if (!fBas.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorBase.getSelectedItem().toString());
            }
            cell = mySheet.getRow(11).getCell(17);
            if (!fBas1.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorBase1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(12).getCell(17);
            if (!fBas2.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorBase2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(10).getCell(18);
            if (!fLob.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorLobby.getSelectedItem().toString());
            }

            cell = mySheet.getRow(11).getCell(18);
            if (!fLob1.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorLobby1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(12).getCell(18);
            if (!fLob2.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorLobby2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(10).getCell(19);
            if (!fHal.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorHall.getSelectedItem().toString());
            }
            cell = mySheet.getRow(11).getCell(19);
            if (!fHal1.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorHall1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(12).getCell(19);
            if (!fHal2.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorHall2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(10).getCell(21);
            if (!fSui.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorSuite.getSelectedItem().toString());
            }
            cell = mySheet.getRow(11).getCell(21);
            if (!fSui1.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorSuite1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(12).getCell(21);
            if (!fSui2.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorSuite2.getSelectedItem().toString());
            }

            cell = mySheet.getRow(10).getCell(22);
            if (!fOth.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorOther.getSelectedItem().toString());
            }

            cell = mySheet.getRow(11).getCell(22);
            if (!fOth1.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorOther1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(12).getCell(22);
            if (!fOth2.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorOther2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(10).getCell(23);
            if (!fOth3.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorOther3.getSelectedItem().toString());
            }
            cell = mySheet.getRow(11).getCell(23);
            if (!fOth4.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorOther4.getSelectedItem().toString());
            }
            cell = mySheet.getRow(12).getCell(23);
            if (!fOth5.matches("Select Option")) {
                cell.setCellValue(spinnerIntFloorOther5.getSelectedItem().toString());
            }
            //FEATURES: PRIMARY
            cell = mySheet.getRow(49).getCell(3);
            if (!ExcaPri.matches("Select Option")) {
                cell.setCellValue(spinnerFeatExcav.getSelectedItem().toString());
            }
            cell = mySheet.getRow(50).getCell(3);
            if (!FoundPrim.matches("Select Option")) {
                cell.setCellValue(spinnerFeatFound.getSelectedItem().toString());
            }
            cell = mySheet.getRow(51).getCell(3);
            if (!ExtFramPrim.matches("Select Option")) {
                cell.setCellValue(spinnerExtFrame.getSelectedItem().toString());
            }

            cell = mySheet.getRow(52).getCell(3);
            if (!IntFramPrim.matches("Select Option")) {
                cell.setCellValue(spinnerIntFrame.getSelectedItem().toString());
            }

            cell = mySheet.getRow(53).getCell(3);
            if (!ExtCladPrim.matches("Select Option")) {
                cell.setCellValue(spinnerExtClad.getSelectedItem().toString());
            }
            cell = mySheet.getRow(55).getCell(3);
            if (!lowFloorprim.matches("Select Option")) {
                cell.setCellValue(spinnerLowFl.getSelectedItem().toString());
            }
            cell = mySheet.getRow(56).getCell(3);
            if (!uppFloorPrim.matches("Select Option")) {
                cell.setCellValue(spinnerUppFl.getSelectedItem().toString());
            }
            cell = mySheet.getRow(57).getCell(3);
            if (!roofStrPrim.matches("Select Option")) {
                cell.setCellValue(spinnerRoofStr.getSelectedItem().toString());
            }
            cell = mySheet.getRow(58).getCell(3);
            if (!roofMatPrim.matches("Select Option")) {
                cell.setCellValue(spinnerRoofMat.getSelectedItem().toString());
            }
            //FEATURES: ADDITIONAL 1
            cell = mySheet.getRow(49).getCell(6);
            if (!ExcaAdd1.matches("Select Option")) {
                cell.setCellValue(spinnerFeatExcavAdd1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(50).getCell(6);
            if (!FoundAdd1.matches("Select Option")) {
                cell.setCellValue(spinnerFeatFoundAdd1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(51).getCell(6);
            if (!ExtFramAdd1.matches("Select Option")) {
                cell.setCellValue(spinnerExtFrameAdd1.getSelectedItem().toString());
            }

            cell = mySheet.getRow(52).getCell(6);
            if (!IntFramAdd1.matches("Select Option")) {
                cell.setCellValue(spinnerIntFrameAdd1.getSelectedItem().toString());
            }

            cell = mySheet.getRow(53).getCell(6);
            if (!ExtCladAdd1.matches("Select Option")) {
                cell.setCellValue(spinnerExtCladAdd1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(55).getCell(6);
            if (!lowFloorAdd1.matches("Select Option")) {
                cell.setCellValue(spinnerLowFlAdd1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(56).getCell(6);
            if (!uppFloorAdd1.matches("Select Option")) {
                cell.setCellValue(spinnerUppFlAdd1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(57).getCell(6);
            if (!roofStrAdd1.matches("Select Option")) {
                cell.setCellValue(spinnerRoofStrAdd1.getSelectedItem().toString());
            }
            cell = mySheet.getRow(58).getCell(6);
            if (!roofMatAdd1.matches("Select Option")) {
                cell.setCellValue(spinnerRoofMatAdd1.getSelectedItem().toString());
            }
            //FEATURES: ADDITIONAL 2
            cell = mySheet.getRow(49).getCell(9);
            if (!ExcaAdd2.matches("Select Option")) {
                cell.setCellValue(spinnerFeatExcavAdd2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(50).getCell(9);
            if (!FoundAdd2.matches("Select Option")) {
                cell.setCellValue(spinnerFeatFoundAdd2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(51).getCell(9);
            if (!ExtFramAdd2.matches("Select Option")) {
                cell.setCellValue(spinnerExtFrameAdd2.getSelectedItem().toString());
            }

            cell = mySheet.getRow(52).getCell(9);
            if (!IntFramAdd2.matches("Select Option")) {
                cell.setCellValue(spinnerIntFrameAdd2.getSelectedItem().toString());
            }

            cell = mySheet.getRow(53).getCell(9);
            if (!ExtCladAdd2.matches("Select Option")) {
                cell.setCellValue(spinnerExtCladAdd2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(55).getCell(9);
            if (!lowFloorAdd2.matches("Select Option")) {
                cell.setCellValue(spinnerLowFlAdd2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(56).getCell(9);
            if (!uppFloorAdd2.matches("Select Option")) {
                cell.setCellValue(spinnerUppFlAdd2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(57).getCell(9);
            if (!roofStrAdd2.matches("Select Option")) {
                cell.setCellValue(spinnerRoofStrAdd2.getSelectedItem().toString());
            }
            cell = mySheet.getRow(58).getCell(9);
            if (!roofMatAdd2.matches("Select Option")) {
                cell.setCellValue(spinnerRoofMatAdd2.getSelectedItem().toString());
            }


            /////////// GENERAL REPORT INFORMATION/////////////////////////////////
            //File #:
            cell = mySheet.getRow(3).getCell(3);
            cell.setCellValue(editFile.getText().toString());
            //Strata #:
            cell = mySheet.getRow(4).getCell(3);
            //File #:
            cell.setCellValue(editStrata.getText().toString());
            // Building Name
            cell = mySheet.getRow(5).getCell(3);
            cell.setCellValue(editBuildingName.getText().toString());
            // Street Name
            cell = mySheet.getRow(6).getCell(3);
            cell.setCellValue(editStAdd.getText().toString());
            // City
            cell = mySheet.getRow(7).getCell(3);
            if (city.matches("Other")) {
                cell.setCellValue(editCity.getText().toString());
            } else {
                cell.setCellValue(spinnercity.getSelectedItem().toString());
            }
            // Province
            cell = mySheet.getRow(8).getCell(3);
            if (province.matches("Other")) {
                cell.setCellValue(editProvince.getText().toString());
            } else {

                cell.setCellValue(spinnerprov.getSelectedItem().toString());
            }
            // Inspection Date:
            cell = mySheet.getRow(9).getCell(3);
            cell.setCellValue(editInspectionDate.getText().toString());
            // effective Date:
            cell = mySheet.getRow(10).getCell(3);
            cell.setCellValue(editEffectiveDate.getText().toString());
            // CIV
            cell = mySheet.getRow(7).getCell(9);
            cell.setCellValue(editCIV.getText().toString());
            // MIV
            cell = mySheet.getRow(8).getCell(9);
            cell.setCellValue(editMIV.getText().toString());
            // SiteContactName
            cell = mySheet.getRow(9).getCell(9);
            cell.setCellValue(editSiteContactName.getText().toString());
            // SiteContact Phone
            cell = mySheet.getRow(10).getCell(9);
            cell.setCellValue(editSiteContactPhone.getText().toString());
            //SiteContactSuite/Buzzer
            cell = mySheet.getRow(11).getCell(9);
            cell.setCellValue(editSiteContactSB.getText().toString());
            // Residential Suites
            cell = mySheet.getRow(30).getCell(1);
            cell.setCellValue(editResiSuitesNo.getText().toString());
            //Commercial Units
            cell = mySheet.getRow(31).getCell(1);
            cell.setCellValue(editCommUnitsNo.getText().toString());
            //Buildings
            cell = mySheet.getRow(31).getCell(4);
            cell.setCellValue(editBuildingNo.getText().toString());
            // Construction Year
            cell = mySheet.getRow(43).getCell(1);
            cell.setCellValue(editConstYear.getText().toString());
            //Zoning
            cell = mySheet.getRow(42).getCell(7);
            cell.setCellValue(editZoning.getText().toString());
            //Parkingsurface
            cell = mySheet.getRow(71).getCell(15);
            cell.setCellValue(editParkSur.getText().toString());
            //UG
            cell = mySheet.getRow(71).getCell(18);
            cell.setCellValue(editParkUG.getText().toString());
            // GRG
            cell = mySheet.getRow(71).getCell(21);
            cell.setCellValue(editParkGRG.getText().toString());
            //CP
            cell = mySheet.getRow(71).getCell(23);
            cell.setCellValue(editParkCP.getText().toString());


            ////////////////////SITE FEATURES////////////////////////////
            //Site Area
            cell = mySheet.getRow(42).getCell(1);
            cell.setCellValue(editSiteArea.getText().toString());

            //Soft%%%
            cell = mySheet.getRow(13).getCell(1);
            cell.setCellValue(editSoftPerc.getText().toString() + "%");
            //Soft Quality
 /*
            //Hard %
            cell = mySheet.getRow(11).getCell(6);
            cell.setCellValue(editHardPerc.getText().toString());
*/

            //SITEFEATURES CHECKBOXES
            cell = mySheet.getRow(16).getCell(2);
            if (grass.isChecked()) {
                cell.setCellValue(true);
            }
            cell = mySheet.getRow(18).getCell(11);
            if (surfacingparking.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(17).getCell(2);
            if (trees.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(18).getCell(2);
            if (scrubs.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(19).getCell(2);
            if (irrigationsys.isChecked()) {
                cell.setCellValue(true);
            }
            //box1
            cell = mySheet.getRow(20).getCell(0);
                cell.setCellValue(box1L);
            cell = mySheet.getRow(20).getCell(2);
            if (!box1L.matches("")) {
                cell.setCellValue(true);
            }

            //box2
            cell = mySheet.getRow(21).getCell(0);
            cell.setCellValue(box2L);
            cell = mySheet.getRow(21).getCell(2);
            if (!box2L.matches("")) {
                cell.setCellValue(true);
            }
            //box3
            cell = mySheet.getRow(22).getCell(0);
            cell.setCellValue(box3L);
            cell = mySheet.getRow(22).getCell(2);
            if (!box3L.matches("")) {
                cell.setCellValue(true);
            }
            //box4
            cell = mySheet.getRow(23).getCell(0);
            cell.setCellValue(box4L);
            cell = mySheet.getRow(23).getCell(2);
            if (!box4L.matches("")) {
                cell.setCellValue(true);
            }
            //box5
            cell = mySheet.getRow(24).getCell(0);
            cell.setCellValue(box5L);
            cell = mySheet.getRow(24).getCell(2);
            if (!box5L.matches("")) {
                cell.setCellValue(true);
            }
            //box6
            cell = mySheet.getRow(25).getCell(0);
            cell.setCellValue(box6L);
            cell = mySheet.getRow(25).getCell(2);
            if (!box6L.matches("")) {
                cell.setCellValue(true);
            }
            //box7
            cell = mySheet.getRow(22).getCell(9);
            cell.setCellValue(box7L);
            cell = mySheet.getRow(22).getCell(11);
            if (!box7L.matches("")) {
                cell.setCellValue(true);
            }
            //box8
            cell = mySheet.getRow(23).getCell(9);
            cell.setCellValue(box8L);
            cell = mySheet.getRow(23).getCell(11);
            if (!box8L.matches("")) {
                cell.setCellValue(true);
            }
            //box9
            cell = mySheet.getRow(24).getCell(9);
            cell.setCellValue(box9L);
            cell = mySheet.getRow(24).getCell(11);
            if (!box9L.matches("")) {
                cell.setCellValue(true);
            }
            //box10
            /*
            cell = mySheet.getRow(25).getCell(9);
            cell.setCellValue(box10L);
            cell = mySheet.getRow(25).getCell(11);
            if ( !box10L.matches("")) {
                cell.setCellValue(true);
            }*/


//
            cell = mySheet.getRow(16).getCell(5);
            if (bballcourt.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(17).getCell(5);
            if (benches.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(18).getCell(5);
            if (bikerack.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(19).getCell(5);
            if (curbing.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(20).getCell(5);
            if (entrancecanopy.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(21).getCell(5);
            if (entranceg.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(22).getCell(5);
            if (fencing.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(23).getCell(5);
            if (garbageshelter.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(24).getCell(5);
            if (gazebo.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(25).getCell(5);
            if (lighting.isChecked()) {
                cell.setCellValue(true);
            }
            //

            cell = mySheet.getRow(16).getCell(8);
            if (mailkiosk.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(17).getCell(8);
            if (metalpoles.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(18).getCell(8);
            if (outdoorspool.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(19).getCell(8);
            if (patios.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(20).getCell(8);
            if (planters.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(21).getCell(8);
            if (playground.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(22).getCell(8);
            if (pedestriangates.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(23).getCell(8);
            if (railing.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(24).getCell(8);
            if (retainingwalls.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(25).getCell(8);
            if (roadways.isChecked()) {
                cell.setCellValue(true);
            }

            //
            cell = mySheet.getRow(16).getCell(11);
            if (signs.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(17).getCell(11);
            if (storageshed.isChecked()) {
                cell.setCellValue(true);
            }
/*              SURFACING MEANT HERE
            cell = mySheet.getRow(18).getCell(11);
            if (.isChecked()) {
                cell.setCellValue(true);
            }
*/

            cell = mySheet.getRow(19).getCell(11);
            if (tenniscourt.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(20).getCell(11);
            if (walkways.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(21).getCell(11);
            if (waterfeature.isChecked()) {
                cell.setCellValue(true);
            }


            ////GENERAL FEATURES
            cell = mySheet.getRow(45).getCell(3);
            cell.setCellValue(numberoffloorsgen.getText().toString());

            cell = mySheet.getRow(38).getCell(9);
            cell.setCellValue(box1PR);
            cell = mySheet.getRow(38).getCell(11);
            if (!box1PR.matches("")) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(39).getCell(9);
            cell.setCellValue(box2PR);
            cell = mySheet.getRow(39).getCell(11);
            if (!box2PR.matches("")) {
                cell.setCellValue(true);
            }
            //STRUCTURES CHECKHOXES
            cell = mySheet.getRow(33).getCell(2);
            if (amenitybuilding.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(34).getCell(2);
            if (amenityroom.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(35).getCell(2);
            if (bikeroom.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(36).getCell(2);
            if (carpports.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(37).getCell(2);
            if (changer.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(38).getCell(2);
            if (commonkitch.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(39).getCell(2);
            if (commonwash.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(40).getCell(2);
            if (crawlspace.isChecked()) {
                cell.setCellValue(true);
            }
//

            cell = mySheet.getRow(33).getCell(5);
            if (exerciser.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(34).getCell(5);
            if (garagedetached.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(35).getCell(5);
            if (garageatt.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(36).getCell(5);
            if (gym.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(37).getCell(5);
            if (indoorhott.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(38).getCell(5);
            if (indoorswimp.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(39).getCell(5);
            if (laundryroom.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(40).getCell(5);
            if (library.isChecked()) {
                cell.setCellValue(true);
            }

            //

            cell = mySheet.getRow(33).getCell(8);
            if (mechroom.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(34).getCell(8);
            if (meetingr.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(35).getCell(8);
            if (multipur.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(36).getCell(8);
            if (office.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(37).getCell(8);
            if (parkadeconc.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(38).getCell(8);
            if (parkadeground.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(39).getCell(8);
            if (partyroom.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(40).getCell(8);
            if (readingroom.isChecked()) {
                cell.setCellValue(true);
            }
//

            cell = mySheet.getRow(33).getCell(11);
            if (roofg.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(34).getCell(11);
            if (sauna.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(35).getCell(11);
            if (steamroom.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(36).getCell(11);
            if (storage.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(37).getCell(11);
            if (workshop.isChecked()) {
                cell.setCellValue(true);
            }
/*
            cell = mySheet.getRow(38).getCell(11);
            if (indoorswimp.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(39).getCell(11);
            if (laundryroom.isChecked()) {
                cell.setCellValue(true);
            }
*/
            cell = mySheet.getRow(40).getCell(11);
            if (andvar.isChecked()) {
                cell.setCellValue(true);
            }else{

                cell.setCellValue(false);
            }

            //extras

            cell = mySheet.getRow(60).getCell(5);
            if (balconies.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(61).getCell(5);
            if (canopies.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(60).getCell(8);
            if (roofdeck.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(61).getCell(8);
            if (roofoverhang.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(60).getCell(11);
            if (skylights.isChecked()) {
                cell.setCellValue(true);
            }

            ////////////////STRUCTURES FEATURE///////////////
            // Excavations

            //Foundations

            //Exterior Frame

            //Interior Frame

            //Exterior Cladding

            //Lowest Floor

            //Upper Floors(s)

            //Roof Structure

            //Roof Material


            ///INTERIOR FINISHES/////////////////
            ////CEILING
            //Parkade
            // Commer
            //Basement
            //Lobby
            //Halls
            //Suite(s)
            //Other
            //Other
            ////WALLS
            //Parkade
            // Commer
            //Basement
            //Lobby
            //Halls
            //Suite(s)
            //Other
            //Other
            ////FLOOR
            //Parkade
            // Commer
            //Basement
            //Lobby
            //Halls
            //Suite(s)
            //Other
            //Other
            ////INTERIOR APPLIANCES///////////////////
            ////APPLIANCES:
            ///Individual Suite Appliances

            cell = mySheet.getRow(15).getCell(16);
            if (dishwasher.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(16).getCell(16);
            if (fridge.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(17).getCell(16);
            if (garburator.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(18).getCell(16);
            if (microwavewhood.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(19).getCell(16);
            if (microwavenohood.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(20).getCell(16);
            if (centralvac.isChecked()) {
                cell.setCellValue(true);
            }

//

            cell = mySheet.getRow(15).getCell(20);
            if (stoveelectr.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(16).getCell(20);
            if (stovegas.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(17).getCell(20);
            if (stovetopelec.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(18).getCell(20);
            if (wallovenelec.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(19).getCell(20);
            if (stovetopgas.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(20).getCell(20);
            if (wallovengas.isChecked()) {
                cell.setCellValue(true);
            }
//

            cell = mySheet.getRow(15).getCell(24);
            if (dryer.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(16).getCell(24);
            if (washer.isChecked()) {
                cell.setCellValue(true);
            }
/*
            cell = mySheet.getRow(17).getCell(25);
            if (stovetopelec.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(18).getCell(25);
            if (wallovenelec.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(19).getCell(25);
            if (stovetopgas.isChecked()) {
                cell.setCellValue(true);
            }
*/


            cell = mySheet.getRow(20).getCell(24);
            if (fireplace.isChecked()) {
                cell.setCellValue(true);
            }

            //COMMON BUILDING APPLIANCES
            cell = mySheet.getRow(29).getCell(14);
            cell.setCellValue(editIntCmnApp_chkbx.getText().toString());

            cell = mySheet.getRow(29).getCell(18);
            cell.setCellValue(editIntCmnApp_chkbx1.getText().toString());
            cell = mySheet.getRow(29).getCell(22);
            cell.setCellValue(editIntCmnApp_chkbx2.getText().toString());




            cell = mySheet.getRow(34).getCell(18);
            cell.setCellValue(editMechElev_chkbx8.getText().toString());




            cell = mySheet.getRow(34).getCell(22);
            cell.setCellValue(editMechElev_chkbx9.getText().toString());



            //SECURITY EDIT TEXTS
            cell = mySheet.getRow(53).getCell(18);
            cell.setCellValue(editSecSysCamNo_chkbx.getText().toString());
            cell = mySheet.getRow(54).getCell(14);
            if(parkadegates.isChecked()){
            cell.setCellValue(editSecSysParkadeNo_chkbx.getText().toString());
            }



            cell = mySheet.getRow(54).getCell(18);
            cell.setCellValue(editSecSysCamBk_chkbx.getText().toString());


            cell = mySheet.getRow(28).getCell(16);
            if (garbagecompactor.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(28).getCell(20);
            if (commonwashers.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(28).getCell(24);
            if (commondryers.isChecked()) {
                cell.setCellValue(true);
            }
            //HVAC SYSTEMS
            cell = mySheet.getRow(37).getCell(16);
            if (aircondiioninghvac.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(38).getCell(16);
            if (ceilingfanshvac.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(39).getCell(16);
            if (elecbaseboard.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(40).getCell(16);
            if (forcedair.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(41).getCell(16);
            if (geothermal.isChecked()) {
                cell.setCellValue(true);
            }

            //

            cell = mySheet.getRow(37).getCell(20);
            if (heatpumphvac.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(38).getCell(20);
            if (hotwaterbaseboard.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(39).getCell(20);
            if (hotwaterradiant.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(40).getCell(20);
            if (hotwaterradiator.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(41).getCell(20);
            if (rooftophvacunits.isChecked()) {
                cell.setCellValue(true);
            }
//
            cell = mySheet.getRow(37).getCell(24);
            if (spaceheaterhvac.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(38).getCell(24);
            if (ventsys.isChecked()) {
                cell.setCellValue(true);
            }

            //HVAC SYSTEMS EDITBOXCHECKBOXES
            //box1
            cell = mySheet.getRow(39).getCell(21);
            cell.setCellValue(box1HVAC);
            cell = mySheet.getRow(39).getCell(24);
            if (!box1HVAC.matches("")) {
                cell.setCellValue(true);
            }

            //box2
            cell = mySheet.getRow(40).getCell(21);
            cell.setCellValue(box2HVAC);
            cell = mySheet.getRow(40).getCell(24);
            if (!box2HVAC.matches("")) {
                cell.setCellValue(true);
            }
            //box3
            cell = mySheet.getRow(41).getCell(21);
            cell.setCellValue(box3HVAC);
            cell = mySheet.getRow(41).getCell(24);
            if (!box3HVAC.matches("")) {
                cell.setCellValue(true);
            }
//ELECTRICAL SYS
            cell = mySheet.getRow(46).getCell(16);
            if (circuitbreakers.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(47).getCell(16);
            if (fuseboxes.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(48).getCell(16);
            if (elecEstim.isChecked()) {
                cell.setCellValue(true);
            }
            //PLUMBING SYS
            cell = mySheet.getRow(46).getCell(20);
            if (copper.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(47).getCell(20);
            if (PEX.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(48).getCell(20);
            if (plumest.isChecked()) {
                cell.setCellValue(true);
            }
            //FIRE PROTECTION
            cell = mySheet.getRow(59).getCell(16);
            if (smokedetwired.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(60).getCell(16);
            if (heatdetwired.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(61).getCell(16);
            if (firepanel.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(62).getCell(16);
            if (firealarms.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(63).getCell(16);
            if (pullstations.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(64).getCell(16);
            if (fireexting.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(65).getCell(16);
            if (emerglights.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(66).getCell(16);
            if (generator.isChecked()) {
                cell.setCellValue(true); }

            //

            cell = mySheet.getRow(59).getCell(20);
            if (smokedetbatt.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(60).getCell(20);
            if (heatdetbatt.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(62).getCell(20);
            if (handicapaccess.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(63).getCell(20);
            if (elevramp.isChecked()) {
                cell.setCellValue(true);
            }


            //FIRE EDITCHECKBOX
            //box1
            cell = mySheet.getRow(64).getCell(17);
            cell.setCellValue(box1FIR);
            cell = mySheet.getRow(64).getCell(20);
            if (!box1FIR.matches("")) {
                cell.setCellValue(true);
            }

            //box2
            cell = mySheet.getRow(65).getCell(17);
            cell.setCellValue(box2FIR);
            cell = mySheet.getRow(65).getCell(20);
            if (!box2FIR.matches("")) {
                cell.setCellValue(true);
            }
            //box3
            cell = mySheet.getRow(66).getCell(17);
            cell.setCellValue(box3FIR);
            cell = mySheet.getRow(66).getCell(20);
            if (!box3FIR.matches("")) {
                cell.setCellValue(true);
            }

            //SPRINKLERS SYSTEM
            cell = mySheet.getRow(59).getCell(24);
            if (nosprink.isChecked()) {
                cell.setCellValue(1);
            }

            cell = mySheet.getRow(59).getCell(24);
            if (sprinklersthroughout.isChecked()) {
                cell.setCellValue(2);
            }
            cell = mySheet.getRow(59).getCell(24);
            if (parkadeboileronly.isChecked()) {
                cell.setCellValue(3);
            }
            cell = mySheet.getRow(59).getCell(24);
            if (firewallsTHs.isChecked()) {
                cell.setCellValue(4);
            }
            //STANDPIPE SYSTEM
            cell = mySheet.getRow(65).getCell(24);
            if (nostandpipe.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(66).getCell(24);
            if (stairwells.isChecked()) {
                cell.setCellValue(true);
            }
            cell = mySheet.getRow(67).getCell(24);
            if (wallcabinets.isChecked()) {
                cell.setCellValue(true);
            }

            ///Common Bulding Appliances
            /// Notes
            /////////MECHANICAL///////

            /////SAFETY AND SECURITY//////////
            cell = mySheet.getRow(51).getCell(16);
            if (prewiredalarm.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(52).getCell(16);
            if (keyscanaccess.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(53).getCell(16);
            if (parkadegates.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(51).getCell(20);
            if (enterphonesystem.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(52).getCell(20);
            if (videosurv.isChecked()) {
                cell.setCellValue(true);
            }

            cell = mySheet.getRow(51).getCell(24);
            if (intercomsystem.isChecked()) {
                cell.setCellValue(true);
            }

            //SECURITYSYS EDITCHECKBOX
            //box1
            cell = mySheet.getRow(52).getCell(21);
            cell.setCellValue(box1SEC);
            cell = mySheet.getRow(52).getCell(24);
            if (!box1SEC.matches("")) {
                cell.setCellValue(true);
            }

            //box2
            cell = mySheet.getRow(53).getCell(21);
            cell.setCellValue(box2SEC);
            cell = mySheet.getRow(53).getCell(24);
            if (!box2SEC.matches("")) {
                cell.setCellValue(true);
            }
            //box3
            cell = mySheet.getRow(54).getCell(21);
            cell.setCellValue(box3SEC);
            cell = mySheet.getRow(54).getCell(24);
            if (!box3SEC.matches("")) {
                cell.setCellValue(true);
            }

            /////NOTES////////////////////
            //INDIVIDUAL SUITE APPLIANCES NOTES
            cell = mySheet.getRow(22).getCell(14);
            cell.setCellValue(editIntIndvNote.getText().toString());
            cell = mySheet.getRow(23).getCell(14);
            cell.setCellValue(editIntIndvNote1.getText().toString());
            //COMMON BUILDING APPLIANCES NOTES
            cell = mySheet.getRow(31).getCell(14);
            cell.setCellValue(editIntCmnAppNote.getText().toString());
            cell = mySheet.getRow(32).getCell(14);
            cell.setCellValue(editIntCmnAppNote1.getText().toString());
            //HVAC SYSTEM NOTES
            cell = mySheet.getRow(42).getCell(14);
            cell.setCellValue(editMechHvacNote.getText().toString());
            cell = mySheet.getRow(43).getCell(14);
            cell.setCellValue(editMechHvacNote2.getText().toString());
            //SECURITY SYSTEM NOTES
            cell = mySheet.getRow(55).getCell(14);
            cell.setCellValue(editSecSysNote.getText().toString());
            cell = mySheet.getRow(56).getCell(14);
            cell.setCellValue(editSecSysNote1.getText().toString());
            //FIRE PROTECTION NOTES
            cell = mySheet.getRow(68).getCell(14);
            cell.setCellValue(editFireProNote.getText().toString());
            cell = mySheet.getRow(69).getCell(14);
            cell.setCellValue(editFireProNote1.getText().toString());
            //ELECTRICAL NOTES
            cell = mySheet.getRow(47).getCell(21);
            cell.setCellValue(editMechElec_chkbx1.getText().toString());
            //MAIN NOTES
            cell = mySheet.getRow(64).getCell(0);
            cell.setCellValue(editNotes.getText().toString());
            cell = mySheet.getRow(66).getCell(0);
            cell.setCellValue(editNotes1.getText().toString());
            cell = mySheet.getRow(68).getCell(0);
            cell.setCellValue(editNotes2.getText().toString());
            cell = mySheet.getRow(70).getCell(0);
            cell.setCellValue(editNotes3.getText().toString());

            cell = mySheet.getRow(72).getCell(0);
            if(garbagecompactor.isChecked()) {
                cell.setCellValue("Garbage Compactor: " + editIntCmnApp_chkbx.getText().toString() + "units, " + spinnerCmnBuildApp.getSelectedItem().toString() );
            }

            cell = mySheet.getRow(72).getCell(3);
            if(commonwashers.isChecked()) {
                cell.setCellValue("Common Washers: " + editIntCmnApp_chkbx1.getText().toString() + "units, " + spinnerCmnBuildApp1.getSelectedItem().toString() );
            }
            cell = mySheet.getRow(72).getCell(6);
            if(commondryers.isChecked()) {
                cell.setCellValue("Common Dryers: " + editIntCmnApp_chkbx2.getText().toString() + "units, " + spinnerCmnBuildApp2.getSelectedItem().toString() );
            }
            cell = mySheet.getRow(72).getCell(9);
            if(Smechelev.matches("Yes")) {
                cell.setCellValue("Elevator: " + editMechElev_chkbx8.getText().toString() + "elevators, serving: " + editMechElev_chkbx9.getText().toString() + " levels" );
            }

            cell = mySheet.getRow(73).getCell(0);
            if(parkadegates.isChecked()) {
                cell.setCellValue("Parkade Gates: " + editSecSysParkadeNo_chkbx.getText().toString() + " gates" );
            }

            cell = mySheet.getRow(73).getCell(3);
            if(videosurv.isChecked()) {
                cell.setCellValue("Video Surveillance: " + editSecSysCamNo_chkbx.getText().toString() + " cameras, backup: " + editSecSysCamBk_chkbx.getText().toString() );
            }

            cell = mySheet.getRow(73).getCell(6);
            if (nosprink.isChecked()) {
                cell.setCellValue("Sprinkler System: No sprinkler");
            }else if(sprinklersthroughout.isChecked()){
                cell.setCellValue("Sprinkler System: Sprinklers throughout");

            }else if(parkadeboileronly.isChecked()){
                cell.setCellValue("Sprinkler System: Parkade/Boiler only");

            }else if (firewallsTHs.isChecked()){
                cell.setCellValue("Sprinkler System: Firewalls (THs)");
            }

            cell = mySheet.getRow(73).getCell(9);
            if(fireplace.isChecked()) {
                cell.setCellValue("Fireplace(s): " + fireplacequan.getText().toString() + " " + fireplacespinner.getSelectedItem().toString() );
            }

            // effective Date:
            cell = mySheet.getRow(72).getCell(1);
            cell.setCellValue("Effective Date: " + editEffectiveDate.getText().toString());


            myInput.close();

            Calendar c = Calendar.getInstance();
            System.out.println("Current time => " + c.getTime());

            SimpleDateFormat df = new SimpleDateFormat("MMM-dd-yy");
            String formattedDate = df.format(c.getTime());

            File reportDirectory = new File(String.valueOf(Environment.getExternalStoragePublicDirectory(
                    Environment.DIRECTORY_DOCUMENTS + "/Reports/" + "Reports_" + formattedDate )));
            reportDirectory.mkdirs();


             if (Strata.matches("")){

                File file = new File(reportDirectory, editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() + "_" + formattedDate.toString() + ".xls");


                    if (file.exists()) {
                        FileOutputStream outF = new FileOutputStream(file, false);
                        myWorkBook.write(outF);
                        outF.close();
                        Toast.makeText(getApplicationContext(), "The file has been overwritten! (" + editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() + "_" + formattedDate.toString() + ".xls)",
                                Toast.LENGTH_LONG).show();

                    } else {
                        FileOutputStream outF = new FileOutputStream(file);
                        myWorkBook.write(outF);
                        outF.close();
                        Toast.makeText(getApplicationContext(), "The file has been created! (" + editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() + "_" + formattedDate.toString() + ".xls)",
                                Toast.LENGTH_LONG).show();
                    }

            }else if (Bname.matches("")){


                File file = new File(reportDirectory, editStrata.getText().toString() + "_F#" + editFile.getText().toString() + "_" + formattedDate.toString() + ".xls");

                if(file.exists()){
                    FileOutputStream outF = new FileOutputStream(file, false) ;
                    myWorkBook.write(outF);
                    outF.close();
                    Toast.makeText(getApplicationContext(), "The file has been overwritten!(" + editStrata.getText().toString() + "_F#" + editFile.getText().toString() + "_" + formattedDate.toString() + ".xls)",
                            Toast.LENGTH_LONG).show();
                }else{
                    FileOutputStream outF = new FileOutputStream(file) ;
                    myWorkBook.write(outF);
                    outF.close();
                    Toast.makeText(getApplicationContext(), "The file has been created!(" + editStrata.getText().toString() + "_F#" + editFile.getText().toString() + "_" + formattedDate.toString() + ".xls)",
                            Toast.LENGTH_LONG).show();}
            }

            else if (editBuildingName != null && editStrata != null) {
                File file = new File(reportDirectory, editStrata.getText().toString() + "_" + editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() + "_" + formattedDate.toString() + ".xls");

                if(file.exists()){
                    FileOutputStream outF = new FileOutputStream(file, false) ;
                    myWorkBook.write(outF);
                    outF.close();
                    Toast.makeText(getApplicationContext(), "The file has been overwritten! (" + editStrata.getText().toString() + "_" + editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() + "_" + formattedDate.toString() + ".xls)"  ,
                            Toast.LENGTH_LONG).show();
                }else{
                    FileOutputStream outF = new FileOutputStream(file) ;
                    myWorkBook.write(outF);
                    outF.close();
                    Toast.makeText(getApplicationContext(), "The file has been created! (" + editStrata.getText().toString() + "_" + editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() + "_" + formattedDate.toString() + ".xls)"  ,
                            Toast.LENGTH_LONG).show();
                }

            }




        }catch (Exception e){e.printStackTrace(); }

        //return;
    }


    public static boolean isExternalStorageReadOnly() {
        String extStorageState = Environment.getExternalStorageState();
        return Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState);
    }

    public static boolean isExternalStorageAvailable() {
        String extStorageState = Environment.getExternalStorageState();
        return Environment.MEDIA_MOUNTED.equals(extStorageState);
    }
/*
    private void initImageLoader() {
        // for universal image loader
        DisplayImageOptions defaultOptions = new DisplayImageOptions.Builder()
                .cacheOnDisc().imageScaleType(ImageScaleType.EXACTLY_STRETCHED)
                .bitmapConfig(Bitmap.Config.RGB_565).build();
        ImageLoaderConfiguration.Builder builder = new ImageLoaderConfiguration.Builder(
                this).defaultDisplayImageOptions(defaultOptions).memoryCache(
                new WeakMemoryCache());

        ImageLoaderConfiguration config = builder.build();
        imageLoader = ImageLoader.getInstance();
        imageLoader.init(config);
    }

    private void init() {

        handler = new Handler();
        gridGalleryy = (GridView) findViewById(R.id.gridGallery);
        gridGalleryy.setFastScrollEnabled(true);
        adapter = new CustomGalleryAdapter(getApplicationContext(), imageLoader);
        adapter.setMultiplePick(false);
        gridGalleryy.setAdapter(adapter);

        gridGalleryy.setOnItemClickListener(new AdapterView.OnItemClickListener() {
            @Override
            public void onItemClick(AdapterView<?> adapterView, View view, int i, long l) {
                Toast.makeText(MainActivity.this,""+imagePaths.get(i),Toast.LENGTH_LONG).show();
            }
        });

        viewSwitcher = (ViewSwitcher) findViewById(R.id.viewSwitcher);
        viewSwitcher.setDisplayedChild(1);

        imgSinglePick = (ImageView) findViewById(R.id.imgSinglePick);
        imgSinglePick.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                if(imagePaths != null)// if minimum image is choose
                    Toast.makeText(MainActivity.this,""+imagePaths.get(0),Toast.LENGTH_LONG).show();
            }
        });

        btnGalleryPick = (Button) findViewById(R.id.btnGalleryPick);
        btnGalleryPick.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                //Open camera intent
                Intent i = new Intent("android.media.action.IMAGE_CAPTURE");
                startActivityForResult(i, 100);
            }
        });

        btnGalleryPickMul = (Button) findViewById(R.id.btnGalleryPickMul);
        btnGalleryPickMul.setOnClickListener(new View.OnClickListener() {

            @Override
            public void onClick(View v) {
                //Manifest recognize our multiple request by this way
                Intent i = new Intent(Action.ACTION_MULTIPLE_PICK);
                startActivityForResult(i, 200);
            }
        });

    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        imagePaths = new ArrayList<String>();
        Calendar c = Calendar.getInstance();
        System.out.println("Current time => " + c.getTime());
        SimpleDateFormat df = new SimpleDateFormat("MMMM-dd-yy");
        String formattedDate = df.format(c.getTime());
        if (requestCode == 200 && resultCode == Activity.RESULT_OK) {
            String[] all_path = data.getStringArrayExtra("all_path");

            ArrayList<CustomGallery> dataT = new ArrayList<CustomGallery>();

            for (String string : all_path) {
                CustomGallery item = new CustomGallery();
                item.sdcardPath = string;
                imagePaths.add(string);
                dataT.add(item);
            }

            viewSwitcher.setDisplayedChild(0);
            adapter.addAll(dataT);

        } else if (requestCode == 100 && resultCode == Activity.RESULT_OK) {
            EditText editFile = (EditText) findViewById(R.id.editFile);
            EditText editStrata = (EditText) findViewById(R.id.editStrata);
            EditText editBuildingName = (EditText) findViewById(R.id.editBuildingName);
            //EditText editSpinnerInspectionBy = (EditText) findViewById(spinnerInspectionBy);
            Bitmap thumbnail = (Bitmap) data.getExtras().get("data");
            ByteArrayOutputStream bytes = new ByteArrayOutputStream();
            thumbnail.compress(Bitmap.CompressFormat.JPEG, 100, bytes);
            if (editStrata != null && editBuildingName != null) {
                File destination = new File(Environment.getExternalStorageDirectory() + "/DCIM/Normac",
                        editStrata.getText().toString() + "_" + editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() + formattedDate.toString() + ".jpg");
                FileOutputStream fo;
                try {
                    destination.createNewFile();
                    fo = new FileOutputStream(destination);
                    fo.write(bytes.toByteArray());
                    fo.close();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }

            } else if (editStrata == null) {
                File destination = new File(Environment.getExternalStorageDirectory() + "/DCIM/Normac",
                        editBuildingName.getText().toString() + "_F#" + editFile.getText().toString() + formattedDate.toString() + ".jpg");
                FileOutputStream fo;
                try {
                    destination.createNewFile();
                    fo = new FileOutputStream(destination);
                    fo.write(bytes.toByteArray());
                    fo.close();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }

            } else if (editBuildingName == null) {
                File destination = new File(Environment.getExternalStorageDirectory() + "/DCIM/Normac",
                        editStrata.getText().toString() + "_F#" + editFile.getText().toString() + formattedDate.toString() + ".jpg");
                FileOutputStream fo;
                try {
                    destination.createNewFile();
                    fo = new FileOutputStream(destination);
                    fo.write(bytes.toByteArray());
                    fo.close();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
*/
}
