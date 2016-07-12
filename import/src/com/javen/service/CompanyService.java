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
 * ��ȡExcel�������е����ݡ���������
 * @author guoxiaoHui
 * @since 2016-7-6 13:28:22
 */
public class CompanyService { 
	/**
	 * ��ȡexcel�ļ�
	 * @param file �ļ�·��
	 * @return ��ȡ�ļ���
	 */
	public static List<Company> getAllByExcel(String file) throws IOException{
		InputStream is = new FileInputStream(file);
		@SuppressWarnings("resource")
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		List<Company> list = new ArrayList<Company>(); 
		// ѭ��������Sheet
		for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if (hssfSheet == null) {
				continue;
			}
			
			// ѭ����Row
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
	 * ������ֵ
	 * @return ���ص�ֵ
	 */
	private static String getValue(HSSFCell hssfCell) {
		if(hssfCell == null){
			return "";
		}
		
		hssfCell.setCellType(1); 	//���õ�Ԫ��ΪString
		// �����ַ������͵�ֵ 
		return String.valueOf(hssfCell.getStringCellValue());
	}
	
	/**
	 * ���ݴ���ļ���������������
	 * @param companies ����ļ���
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
			System.out.println("���ݲ���ɹ�");
		} catch (Exception e) {
			count = 0;
			e.printStackTrace();
			System.err.println("���ݲ���ʧ��");
		}finally {
			bhepler.closeConnection(connection, ps, rs);
			System.out.println("�ر�����");
		}
		
		return count;
	}
}
