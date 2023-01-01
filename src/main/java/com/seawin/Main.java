package com.seawin;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class Main {

    private static HashMap<String,String> addressMap;

    public static void main(String[] args) throws Exception{
        //createWareHouseBill();
        addressMap = readAmazonAddress();
        String date = "8/1/2023";
        String billOfLading = "EGLV143258154421";

        /*String containerNumber = "TRHU5234335";
        String boxBillOriginalFileName = "TRHU5234335+EGLV143258154421拆柜资料.xlsx";*/
/*
        String containerNumber = "MATU2716900";
        String boxBillOriginalFileName = "MATU2716900+装箱单-确认.xlsx";*/
        String containerNumber = "MATU5207017";
        String boxBillOriginalFileName = "MATU5207017+装箱单-确认.xlsx";




        HashMap<String,List<List<String>>> datas = readBoxBill(boxBillOriginalFileName);
        generateWareHouseBill(date,billOfLading,containerNumber,datas);

    }

    private static HashMap<String,String> readAmazonAddress() throws Exception{
        String fileLocation = "E:\\seawin\\seawin-tool\\demo\\amazon-warehouse.xlsx";
        HashMap<String,String> addressMap = new HashMap<>();
        try(FileInputStream file = new FileInputStream(new File(fileLocation));
            Workbook workbook = new XSSFWorkbook(file);){
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                String detail = row.getCell(1).getStringCellValue()+" "+ row.getCell(2).getStringCellValue()+" "+ row.getCell(3).getStringCellValue();
                CellType cellType = row.getCell(4).getCellType();
                if(cellType.equals(CellType.NUMERIC)){
                    detail += " "+Double.valueOf(row.getCell(4).getNumericCellValue()).intValue();
                }else if(cellType.equals(CellType.STRING)){
                    detail += " "+row.getCell(4).getStringCellValue();
                }
                addressMap.put(row.getCell(0).getStringCellValue(),detail);
            }
        }
        System.out.println("addressMap size:"+addressMap.size());
        return addressMap;
    }
    private static HashMap<String,List<List<String>>> readBoxBill(String boxBillOriginalFileName) throws Exception{
        String fileLocation = "E:\\seawin\\seawin-tool\\demo\\"+boxBillOriginalFileName;
        FileInputStream file = new FileInputStream(new File(fileLocation));
        Workbook workbook = new XSSFWorkbook(file);

        Sheet sheet = workbook.getSheetAt(0);

        Set<String> wareHouseSet = new TreeSet<>();
        HashMap<String,List<List<String>>> datas = new LinkedHashMap();

        Map<Integer, List<String>> data = new HashMap<>();
        int i = 0;
        String previousWareHouse = "";
        for (Row row : sheet) {

            String bizCode =  row.getCell(0).getRichStringCellValue().toString();
            if(bizCode.trim().length()==0){
                i++;
                continue;
            }
            if(i>=3){
                String wareHouse = row.getCell(6).getRichStringCellValue().toString();
                if(wareHouse.trim().equalsIgnoreCase("UPS")){
                    i++;
                    continue;
                }
                if(wareHouse.trim().length()==0){
                    wareHouse = previousWareHouse;
                }else{
                    previousWareHouse = wareHouse;
                }
                List<List<String>> wareHouseDataList;
                if(datas.containsKey(wareHouse)){
                    wareHouseDataList = datas.get(wareHouse);
                }else{
                    wareHouseDataList = new LinkedList<>();
                    datas.put(wareHouse,wareHouseDataList);
                }
                double ctn = row.getCell(3).getNumericCellValue();
                double kgs  = row.getCell(4).getNumericCellValue();
                wareHouseDataList.add(List.of(
                        row.getCell(0).getRichStringCellValue().toString(),
                        row.getCell(1).getRichStringCellValue().toString(),
                        row.getCell(2).getRichStringCellValue().toString(),
                        ctn==0d?"":String.valueOf(Double.valueOf(ctn).intValue()),
                        kgs==0d?"":String.valueOf(kgs),
                        row.getCell(8).getRichStringCellValue().toString()));


            }
            i++;
        }

        System.out.println(datas);
        return datas;
    }

    private  static void generateWareHouseBill(
            String date,
            String billOfLading,
            String containerNumber,
            HashMap<String,List<List<String>>> datas) throws Exception{
        Set<String> warehouseSet = datas.keySet();
        for(String warehouse:warehouseSet){
            List<List<String>> wareHouseDataList = datas.get(warehouse);
           /* if(wareHouseDataList.size() == 1){

            }*/
            String shipToName = warehouse;
            String shipToAddress = addressMap.getOrDefault(shipToName,"");
            if(shipToAddress.length()==0){
                System.out.println("[shipToName:"+shipToName+"] not find");
            }
            try{
                createWareHouseBill(date,billOfLading,containerNumber,shipToName,shipToAddress,wareHouseDataList);
            }catch (Exception e){
                e.printStackTrace();
            }

        }
    }

    private static void createWareHouseBill(String date,
                                            String billOfLading,
                                            String containerNumber,
                                            String shipToName,
                                            String shipToAddress,
                                            List<List<String>> shipmentDetailsList)throws Exception{
        String fileLocation = "E:\\seawin\\seawin-tool\\demo\\warehouse-template.xlsx";
        String fileName = shipToName;
        /*
        if(shipToName.equals("私人地址")){
            fileName = shipToName+"-"+shipmentDetails.get(0);
            if(shipmentDetails.size()>=6){
                shipToAddress = shipmentDetails.get(5);
            }
        }*/
        File fileDir = new File("E:\\seawin\\seawin-tool\\demo\\"+containerNumber);
        if(!fileDir.exists()){
            fileDir.mkdirs();
        }
        File targetFile = new File(fileDir,fileName+".xls");
        try (FileInputStream file = new FileInputStream(new File(fileLocation));
             Workbook workbook = new XSSFWorkbook(file);

             FileOutputStream outputStream = new FileOutputStream(targetFile)){

            Sheet sheet = workbook.getSheetAt(0);
            int i = 0;
            int shipmentDetailsCount =0;
            for (Row row : sheet) {
                if(i==3){
                    int j=0;
                    for (Cell cell : row) {
                        if(j==4){
                            cell.setCellValue(date);
                            break;
                        }
                        j++;
                    }
                }
                if(i==4){
                    int j=0;
                    for (Cell cell : row) {
                        if(j==3){
                            cell.setCellValue(cell.getRichStringCellValue().toString()+billOfLading);
                            break;
                        }
                        j++;
                    }
                }
                if(i==5){
                    int j=0;
                    for (Cell cell : row) {
                        if(j==5){
                            cell.setCellValue(containerNumber);
                            break;
                        }
                        j++;
                    }
                }
                if(i==9){
                    for (Cell cell : row) {
                        String shipToNameOriginal = cell.getRichStringCellValue().toString();
                        cell.setCellValue(shipToNameOriginal+shipToName);
                        break;
                    }
                }
                if(i==10){
                    for (Cell cell : row) {
                        String shipToNameOriginal = cell.getRichStringCellValue().toString();
                        cell.setCellValue(shipToNameOriginal+"            "+shipToAddress);
                        break;
                    }
                }


                if(i>=21){
                    if(shipmentDetailsCount<shipmentDetailsList.size()){
                        List<String> shipmentDetails = shipmentDetailsList.get(shipmentDetailsCount);
                        row.getCell(1).setCellValue(shipmentDetails.get(0));
                        row.getCell(2).setCellValue(shipmentDetails.get(1));
                        row.getCell(3).setCellValue(shipmentDetails.get(2));
                        row.getCell(4).setCellValue(shipmentDetails.get(3));
                        row.getCell(5).setCellValue(shipmentDetails.get(4));
                        shipmentDetailsCount++;
                    }
                }
                i++;
            }
            /*
            if(shipmentDetailsList.size()>1){
                sheet.addMergedRegionUnsafe(CellRangeAddress.valueOf("F22:F"+(22+shipmentDetailsList.size()-1)));
            }*/

            workbook.write(outputStream);
            //workbook.close();
            //outputStream.close();
        }

    }



    private static void createWareHouseBill() throws Exception{
        String date = "8/1/2023";
        String billOfLading = "EGLV143258154421";
        String containerNumber = "TRHU5234335";
        String shipToName = "SBD1";
        String shipToAddress = "3388 S Cactus Ave BLOOMINGTON CA 92316-3819";
        List<String> shipmentDetails = List.of("YH002369147","FBA16ZSN1VJW","4JAL8Y5T","52","1072");
        createWareHouseBill(date,billOfLading,containerNumber,shipToName,shipToAddress,List.of(shipmentDetails));
    }


}
