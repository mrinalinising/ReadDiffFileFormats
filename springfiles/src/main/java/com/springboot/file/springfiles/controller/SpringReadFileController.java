package com.springboot.file.springfiles.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.springboot.file.springfiles.User;
import com.springboot.file.springfiles.service.SpringReadFileService;

@Controller
public class SpringReadFileController {
	public static String uploadDirectory = System.getProperty("user.dir")+"/uploads";

	@Autowired
	private SpringReadFileService springReadFileService;
	@Autowired
	private ServletContext context;
	
	@GetMapping(value="/")
	public String home(Model model) {
		model.addAttribute("user",new User());
		List<User> users=springReadFileService.findAll();
		model.addAttribute("users",users);
		return "view/users";
	}
	
	@PostMapping(value="/fileUpload")
	public String uploadFile(@ModelAttribute User user,RedirectAttributes redirectAttributes) {
		
		
		boolean isFlag=springReadFileService.saveDataFromUploadFile(user.getFile());
		if(isFlag) {
			redirectAttributes.addFlashAttribute("successmessage", "File uploaded Successfully!");
		}else {
			redirectAttributes.addFlashAttribute("errormessage", "File upload not done.Please try again!");
		}
		return "redirect:/";
	}
	
	@GetMapping(value="/createExcel")
	public void createExcel(HttpServletRequest req,HttpServletResponse res) {
		System.out.println("================================++++++++++++");
		List<User> users=springReadFileService.findAll();
		boolean flag=springReadFileService.createExcel(users,context,req,res);
		if(flag) {
			String fullPath=req.getServletContext().getRealPath("/resources/reports/"+"employees"+".xls");
			fileDownload(fullPath,res,"employees.xls");
		}
	}

	private void fileDownload(String fullPath, HttpServletResponse res, String fileName) {
		File file=new File(fullPath);
		final int BUFFER_SIZE=4096;
		if(file.exists()) {
			try {
				FileInputStream inputStream=new FileInputStream(file);
				String mimeType=context.getMimeType(fullPath);
				res.setContentType(mimeType);
				res.setHeader("content-disposition", "attachment ; filename ="+fileName);
				OutputStream outputStream=res.getOutputStream();
				byte[] buffer=new byte[BUFFER_SIZE];
				int bytesRead=-1;
				while((bytesRead=inputStream.read(buffer))!=-1) {
					outputStream.write(buffer,0,bytesRead);					
				}
				inputStream.close();
				outputStream.close();
				file.delete();
			}catch(Exception e) {
				e.printStackTrace();
			}
			
		}
	}
	
	

}
