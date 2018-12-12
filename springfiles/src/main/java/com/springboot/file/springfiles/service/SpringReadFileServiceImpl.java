package com.springboot.file.springfiles.service;


import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.h2.result.Row;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.springboot.file.springfiles.User;
import com.springboot.file.springfiles.repository.SpringReadFileRepository;

@Service
@Transactional	
public class SpringReadFileServiceImpl implements SpringReadFileService{
	@Autowired
	private SpringReadFileRepository springReadFileRepository;

	@Override
	public List<User> findAll() {
		return (List<User>) springReadFileRepository.findAll();
	}

	@Override
	public boolean saveDataFromUploadFile(MultipartFile file) {
		boolean isFlag=false;
		String extension=FilenameUtils.getExtension(file.getOriginalFilename());
		if(extension.equalsIgnoreCase("json")) {
			isFlag=readDataFromJson(file);
		}else if(extension.equalsIgnoreCase("csv")) {
			isFlag=readDataFromCsv(file);
		}else if(extension.equalsIgnoreCase("xls")||extension.equalsIgnoreCase("xlsx")) {
			isFlag=readDataFromExcel(file);
		}
		return isFlag;
	}
		
	/*public boolean saveDataFromUploadFile(MultipartFile[] files) {
	boolean isFlag=false;
	String extension=null;
	for(MultipartFile file:files) {
	 extension=FilenameUtils.getExtension(file.getOriginalFilename());
	
	if(extension.equalsIgnoreCase("json")) {
		isFlag=readDataFromJson(file);
	}
    if(extension.equalsIgnoreCase("csv")) {
		isFlag=readDataFromCsv(file);
    }
	if(extension.equalsIgnoreCase("xls")||extension.equalsIgnoreCase("xlsx")) {
		isFlag=readDataFromExcel(file);
	}
   }
	return isFlag;
}*/

	private boolean readDataFromExcel(MultipartFile file) {
		Workbook workbook=getWorkBook(file);
		Sheet sheet=workbook.getSheetAt(0);
		Iterator<org.apache.poi.ss.usermodel.Row> rows=sheet.iterator();
		rows.next();
		while(rows.hasNext()) {
			org.apache.poi.ss.usermodel.Row row=rows.next();
			User user=new User();
			if(row.getCell(0).getCellType()==Cell.CELL_TYPE_STRING) {
				user.setFirstName(row.getCell(0).getStringCellValue());
			}
			if(row.getCell(1).getCellType()==Cell.CELL_TYPE_STRING) {
				user.setLastName(row.getCell(1).getStringCellValue());
			}
			if(row.getCell(2).getCellType()==Cell.CELL_TYPE_STRING) {
				user.setEmail(row.getCell(2).getStringCellValue());
			}
			if(row.getCell(3).getCellType()==Cell.CELL_TYPE_NUMERIC) {
				String phoneNumber=NumberToTextConverter.toText(row.getCell(3).getNumericCellValue());
				user.setPhoneNumber(phoneNumber);
			}else if(row.getCell(3).getCellType()==Cell.CELL_TYPE_STRING) {
				user.setPhoneNumber(row.getCell(3).getStringCellValue());
			}
			user.setFileType(FilenameUtils.getExtension(file.getOriginalFilename()));
			springReadFileRepository.save(user);
			
		}
		return true;
	}

	private Workbook getWorkBook(MultipartFile file) {
		Workbook workbook=null;
		String extension=FilenameUtils.getExtension(file.getOriginalFilename());
		try {
		/*if(extension.equalsIgnoreCase("xlsx")){
			workbook=new XSSFWorkbook(file.getInputStream());
		}else*/ if(extension.equalsIgnoreCase("xls")) {
			workbook= new HSSFWorkbook(file.getInputStream());
		}
	}catch(Exception e) {
		e.printStackTrace(); 
	}

		
		
		return workbook;
	}

	private boolean readDataFromCsv(MultipartFile file) {
		try {
			InputStreamReader reader=new InputStreamReader(file.getInputStream());
			CSVReader csvReader=new CSVReaderBuilder(reader).withSkipLines(1).build();
			List<String[]> rows=csvReader.readAll();
			for(String[] row:rows) {
				springReadFileRepository.save(new User(row[0],row[1],row[2],row[3],FilenameUtils.getExtension(file.getOriginalFilename())));
			}
			return true;
	
		}catch(Exception e) {
		   return false;
		}
	}

	private boolean readDataFromJson(MultipartFile file) {
		try {
			InputStream inputStream=file.getInputStream();
			ObjectMapper mapper=new ObjectMapper();
			List<User> users=Arrays.asList(mapper.readValue(inputStream, User[].class));
			if(users!=null && users.size()>0) {
				for(User user:users) {
					user.setFileType(FilenameUtils.getExtension(file.getOriginalFilename()));
					springReadFileRepository.save(user);
				}
			}
			return true;
		}catch(Exception e) {
		return false;
		}
	}
	
	

	@Override
	public boolean createExcel(List<User> users, ServletContext context, HttpServletRequest req,
			HttpServletResponse res) {
		System.out.println("==========================createExcel++++++++ServiceImpl======");
		String filePath=context.getRealPath("/resources/reports");
		File file=new File(filePath);
		boolean exists=new File(filePath).exists();
		if(!exists) {
			new File(filePath).mkdirs();			
		}
		
		try {
			FileOutputStream outputStream=new FileOutputStream(file+"/"+"employees"+".xls");
			HSSFWorkbook workbook=new HSSFWorkbook();
			HSSFSheet workSheet=workbook.createSheet("Employees");
			workSheet.setDefaultColumnWidth(30);
			
			HSSFCellStyle headerCellStyle=workbook.createCellStyle();
			headerCellStyle.setFillBackgroundColor(HSSFColor.BLUE.index);
			headerCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			
			HSSFRow headerRow=workSheet.createRow(0);
			
			HSSFCell firstName=headerRow.createCell(0);
			firstName.setCellValue("First Name");
			firstName.setCellStyle(headerCellStyle);
			
			HSSFCell lastName=headerRow.createCell(1);
			lastName.setCellValue("Last Name");
			lastName.setCellStyle(headerCellStyle);
			
			HSSFCell email=headerRow.createCell(2);
			email.setCellValue("Email");
			email.setCellStyle(headerCellStyle);
			
			HSSFCell phoneNumber=headerRow.createCell(3);
			phoneNumber.setCellValue("Phone Number");
			phoneNumber.setCellStyle(headerCellStyle);
			
			HSSFCell fileType=headerRow.createCell(4);
			fileType.setCellValue("File Type");
			fileType.setCellStyle(headerCellStyle);
			
			int i=1;
			for(User user:users) {
				HSSFRow bodyRow=workSheet.createRow(i);
				
				HSSFCellStyle bodyCellStyle=workbook.createCellStyle();
				bodyCellStyle.setFillForegroundColor(HSSFColor.WHITE.index);
				
				HSSFCell firstNameValue=bodyRow.createCell(0);
				firstNameValue.setCellValue(user.getFirstName());
				firstNameValue.setCellStyle(bodyCellStyle);
				
				HSSFCell lastNameValue=bodyRow.createCell(1);
				lastNameValue.setCellValue(user.getLastName());
				lastNameValue.setCellStyle(bodyCellStyle);
				
				HSSFCell emailValue=bodyRow.createCell(2);
				emailValue.setCellValue(user.getEmail());
				emailValue.setCellStyle(bodyCellStyle);
				
				HSSFCell phoneNumberValue=bodyRow.createCell(3);
				phoneNumberValue.setCellValue(user.getPhoneNumber());
				phoneNumberValue.setCellStyle(bodyCellStyle);
				
				HSSFCell fileTypeValue=bodyRow.createCell(4);
				fileTypeValue.setCellValue(user.getFileType());
				fileTypeValue.setCellStyle(bodyCellStyle);
				
				i++;
			}
			
			workbook.write(outputStream);
			outputStream.flush();
			outputStream.close();
			return true;


		}catch(Exception e) {
			return false;
		}
	}

}
