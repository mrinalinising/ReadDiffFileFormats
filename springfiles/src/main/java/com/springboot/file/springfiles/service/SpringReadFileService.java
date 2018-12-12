package com.springboot.file.springfiles.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.multipart.MultipartFile;

import com.springboot.file.springfiles.User;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;


public interface SpringReadFileService {

	List<User> findAll();

	boolean saveDataFromUploadFile(MultipartFile file);

	boolean createExcel(List<User> users, ServletContext context, HttpServletRequest req, HttpServletResponse res);
	

}
