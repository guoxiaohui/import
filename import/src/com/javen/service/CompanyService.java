package com.javen.service;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.javen.db.DBhepler;
import com.javen.entity.Company;
/**
 * 读取Excel表中所有的数据、操作数据
 * @author guoxiaoHui
 * @since 2016-7-6 13:28:22
 */
public class CompanyService { 
	/**
	 * 读取excel文件
	 * @param file 文件路径
	 * @return 读取的集合
	 */
	public static List<Company> getAllByExcel(String file) throws IOException{
		InputStream is = new FileInputStream(file);
		@SuppressWarnings("resource")
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		List<Company> list = new ArrayList<Company>(); 
		// 循环工作表Sheet
		for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if (hssfSheet == null) {
				continue;
			}
			
			// 循环行Row
			for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
				if (hssfRow != null) {
					HSSFCell name = hssfRow.getCell(1);
					HSSFCell taxCode = hssfRow.getCell(7);
					if((null == name || name.equals("")) || (null == taxCode || taxCode.equals(""))){
						continue;
					}
					
					HSSFCell phone = hssfRow.getCell(9);
					HSSFCell address = hssfRow.getCell(10);
					int random = new Random().nextInt(89999)+10000;
					long uuid = new Date().getTime()-1300000000000L + rowNum + random;
					list.add(new Company(getValue(name) , getValue(taxCode) , getValue(address) , getValue(phone) , String.valueOf(uuid)));
				}
			}
			
		}
		
		return list;
	}
	
	/**
	 * 处理返回值
	 * @return 返回的值
	 */
	private static String getValue(HSSFCell hssfCell) {
		if(hssfCell == null){
			return "";
		}
		
		hssfCell.setCellType(1); 	//设置单元格为String
		// 返回字符串类型的值 
		return String.valueOf(hssfCell.getStringCellValue());
	}
	
	/**
	 * 根据传入的集合批量保存数据
	 * @param companies 传入的集合
	 * @return 
	 */
	public static int batchSave(List<Company> companies){
		DBhepler bhepler = new DBhepler();
		Connection connection = null;
		PreparedStatement ps = null;
		ResultSet rs = null;
		int count = 1;
		try {
			connection = bhepler.getConnection();
			String sql = "INSERT INTO COMPANY (NAME,COMPANY_NO,TAXCODE,ADDRESS,PHONE) values (?, ?, ? ,? ,?)";
			ps = connection.prepareStatement(sql);
			final int batchSize = 1000;
			for (Company company: companies) {
			    ps.setString(1, company.getName());
			    ps.setString(2, company.getCompanyNo());
			    ps.setString(3, company.getTaxCode());
			    ps.setString(4, company.getAddress());
			    ps.setString(5, company.getPhone());
			    ps.addBatch();
			    if(++count % batchSize == 0) {
			        ps.executeBatch();
			    }
			}
			
			ps.executeBatch();
			System.out.println("数据插入成功");
		} catch (Exception e) {
			count = 0;
			e.printStackTrace();
			System.err.println("数据插入失败");
		}finally {
			bhepler.closeConnection(connection, ps, rs);
			System.out.println("关闭连接");
		}
		
		return count;
	}
}
