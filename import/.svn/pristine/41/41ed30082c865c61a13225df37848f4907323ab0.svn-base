package com.javen.excel;

import java.io.IOException;
import java.util.List;

import com.javen.entity.Company;
import com.javen.service.CompanyService;

public class TestExcelToDb {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		int[] count = new int[12];
		for(int i = 1 ; i <= 12 ; i ++){
			String file = "T:\\excel2003\\excel"+ i +".xls";
			List<Company> list;
			try {
				list = CompanyService.getAllByExcel(file);
				count[i - 1] = CompanyService.batchSave(list);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		System.out.println(count.toString());
		
		
	}

}
