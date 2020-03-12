package com.revincx;
import java.util.*;

import com.google.gson.Gson;

import jxl.Workbook;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.io.*;
//import com.google.gson.*;

public class FetchLK
{
	static String url_info = "http://www.wylkyj.com/ExamAPI/Exam/ScoreQuery/GetStuInfo";
	static String url_score = "http://www.wylkyj.com/ExamAPI/Exam/ScoreQuery/GetStuScore";
	public static void main(String[] args)
	{
		Scanner scanner = new Scanner(System.in);
		System.out.printf("考生号前缀：");
		String prefix = scanner.next();
		//输入考生号的前六位
		scanner.close();
		String stu_no = "";
		File xls = new File("/storage/emulated/0/AppProjects/GradeFetcher/output/g.xls");
		//创建表格文件
		if (!xls.exists()) 
		{
			try
			{
				xls.createNewFile();
				//创建新的表格文件
			} 
			catch (IOException e) 
			{
				e.printStackTrace();
			}
		}
		try 
		{
			WritableWorkbook workbook = Workbook.createWorkbook(xls);
			//创建工作表
			WritableSheet sheet = workbook.createSheet("lk", 0);
			//创建工作簿
			int row = 1;
			//定义初始行号
			while (true)
			//遍历考号
			{
				{
					stu_no = prefix + append(row) ;
					Result result = fetchResult(stu_no, url_score);
					//获取返回的结果
					System.out.println("已获取" + stu_no);
					Grade[] grades = result.datas;
					int subject_num = grades.length;
					if (row == 865) 
					//判断该考生是否不存在
					{
						break;
					}
					int column = 0;
					jxl.write.Label label_no = new jxl.write.Label(column, row, grades[0].stu_no);
					column++;
					jxl.write.Label label_name = new jxl.write.Label(column, row, grades[0].stu_name);
					column++;
					//向单元格写入考生号和考生姓名
					sheet.addCell(label_name);
					sheet.addCell(label_no);
					//把单元格加入工作簿
					int sch_rank = -1;
					while (column < subject_num + 2)
					//遍历学科序号
					{
						String score = Integer.toString((int) grades[column-2].score);
						if(score == null)
						{
							score = "";
						}
						
						int write_column = -1;
						switch(grades[column-2].esub_name)
						{
							case "语文" :
								write_column = 2;
								break;
							case "理科数学" :
								write_column = 3;
								break;
							case "英语" :
								write_column = 4;
								break;
							case "英语虚拟" :
								write_column = 5;
								break;
							case "理科综合" :
								write_column = 6;
								break;
							case "物理" :
								write_column = 7;
								break;
							case "化学" :
								write_column = 8;
								break;
							case "生物" :
								write_column = 9;
								break;
							case "理科总分" :
								write_column = 10;
								sch_rank = grades[column-2].sch_rank;
								break;
						}
						jxl.write.Label label_score = new jxl.write.Label(write_column, row,score);
						sheet.addCell(label_score);
//						String score = Integer.toString((int) grades[column-2].score);
//						if(score == null)
//						{
//							score = "";
//						}
//						jxl.write.Label label_score = new jxl.write.Label(column, row,score);
//						sheet.addCell(label_score);
//						//向工作簿写入考生成绩
						column++;
					}
					jxl.write.Label label_rank = new jxl.write.Label(11,row,Integer.toString(sch_rank));
					sheet.addCell(label_rank);
					row++;
					//使行号加一，移动到下一行
					System.out.println("已添加" + stu_no + grades[0].stu_name);
				}

			}
			System.out.println("写入文件中...");
			workbook.write();
			//把工作表写入到文件
			workbook.close();
			//关闭工作表
			System.out.println("文件已写入");
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}

//		System.out.println(fetchJson(stu_no, url_score));
	}

	static String append(int i)
	//把序号转换为带零的字符串
	{
		String str = Integer.toString(i);
		if (str.length() == 1)
		{
			str = "00" + str;
		}
		else if (str.length() == 2)
		{
			str = "0" + str;
		}
		return str;
	}

	static Result fetchResult(String stu_no, String url)
	//从服务器获取JSON字符串
	{
		String content = "exam_no=14360&e_dbname=exam_14360&stu_no=" + stu_no;
		String json = doPost(url, content);
		Gson gson = new Gson();
		Result result = gson.fromJson(json, Result.class);
		return result;

	}

	public static String doGet(String httpurl)
	{
        HttpURLConnection connection = null;
        InputStream is = null;
        BufferedReader br = null;
        String result = null;// 返回结果字符串
        try
		{
            // 创建远程url连接对象
            URL url = new URL(httpurl);
            // 通过远程url连接对象打开一个连接，强转成httpURLConnection类
            connection = (HttpURLConnection) url.openConnection();
            // 设置连接方式：get
            connection.setRequestMethod("GET");
            // 设置连接主机服务器的超时时间：15000毫秒
            connection.setConnectTimeout(15000);
            // 设置读取远程返回的数据时间：60000毫秒
            connection.setReadTimeout(60000);
            // 发送请求
            connection.connect();
            // 通过connection连接，获取输入流
            if (connection.getResponseCode() == 200)
			{
                is = connection.getInputStream();
                // 封装输入流is，并指定字符集
                br = new BufferedReader(new InputStreamReader(is, "UTF-8"));
                // 存放数据
                StringBuffer sbf = new StringBuffer();
                String temp = null;
                while ((temp = br.readLine()) != null)
				{
                    sbf.append(temp);
                    sbf.append("\r\n");
                }
                result = sbf.toString();
            }
        }
		catch (MalformedURLException e)
		{
            e.printStackTrace();
        }
		catch (IOException e)
		{
            e.printStackTrace();
        }
		finally
		{
            // 关闭资源
            if (null != br)
			{
                try
				{
                    br.close();
                }
				catch (IOException e)
				{
                    e.printStackTrace();
                }
            }

            if (null != is)
			{
                try
				{
                    is.close();
                }
				catch (IOException e)
				{
                    e.printStackTrace();
                }
            }

            connection.disconnect();// 关闭远程连接
        }

        return result;
    }

	public static String doPost(String httpUrl, String param)
	{

        HttpURLConnection connection = null;
        InputStream is = null;
        OutputStream os = null;
        BufferedReader br = null;
        String result = null;
        try
		{
            URL url = new URL(httpUrl);
            // 通过远程url连接对象打开连接
            connection = (HttpURLConnection) url.openConnection();
            // 设置连接请求方式
            connection.setRequestMethod("POST");
            // 设置连接主机服务器超时时间：15000毫秒
            connection.setConnectTimeout(15000);
            // 设置读取主机服务器返回数据超时时间：60000毫秒
            connection.setReadTimeout(60000);

            // 默认值为：false，当向远程服务器传送数据/写数据时，需要设置为true
            connection.setDoOutput(true);
            // 默认值为：true，当前向远程服务读取数据时，设置为true，该参数可有可无
            connection.setDoInput(true);
            // 设置传入参数的格式:请求参数应该是 name1=value1&name2=value2 的形式。
            connection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
            // 设置鉴权信息：Authorization: Bearer da3efcbf-0845-4fe3-8aba-ee040be542c0
            connection.setRequestProperty("Authorization", "Bearer da3efcbf-0845-4fe3-8aba-ee040be542c0");
            // 通过连接对象获取一个输出流
            os = connection.getOutputStream();
            // 通过输出流对象将参数写出去/传输出去,它是通过字节数组写出的
            os.write(param.getBytes());
            // 通过连接对象获取一个输入流，向远程读取
            if (connection.getResponseCode() == 200)
			{

                is = connection.getInputStream();
                // 对输入流对象进行包装:charset根据工作项目组的要求来设置
                br = new BufferedReader(new InputStreamReader(is, "UTF-8"));

                StringBuffer sbf = new StringBuffer();
                String temp = null;
                // 循环遍历一行一行读取数据
                while ((temp = br.readLine()) != null)
				{
                    sbf.append(temp);
                    sbf.append("\r\n");
                }
                result = sbf.toString();
            }
        }
		catch (MalformedURLException e)
		{
            e.printStackTrace();
        }
		catch (IOException e)
		{
            e.printStackTrace();
        }
		finally
		{
            // 关闭资源
            if (null != br)
			{
                try
				{
                    br.close();
                }
				catch (IOException e)
				{
                    e.printStackTrace();
                }
            }
            if (null != os)
			{
                try
				{
                    os.close();
                }
				catch (IOException e)
				{
                    e.printStackTrace();
                }
            }
            if (null != is)
			{
                try
				{
                    is.close();
                }
				catch (IOException e)
				{
                    e.printStackTrace();
                }
            }
            // 断开与远程地址url的连接
            connection.disconnect();
        }
        return result;
    }
}
