
import java.io.File;
import java.io.FileOutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.text.*;
import java.util.Date;
import java.util.Properties;
import java.util.TimeZone;
import java.util.*;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class TestAgentScheduleReport {
	
	 

	public static void main(String[] args)  throws Exception
	{
		String sessionID = null;
		PropertiesLoader loader = new PropertiesLoader();
		
		//Getting time 12 hours ago in PST time zone 
		SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMddHHmmss"); 
		formatter.setTimeZone(TimeZone.getTimeZone("PST"));
		Date strDate = new Date(System.currentTimeMillis()-19*60*60*1000);  
		String strTime=formatter.format(strDate)+"000000";
		
		
		
		String agent = loader.getValue("AGENT");

		
		//apiString = apiString.replace("{strTime}", strTime);
		//System.out.println("api endpoint from props ="+apiString);
		//apiString ="http://m1tida04.mnao.net:8080/api/tes-6-prod/JobRun.getList?query=(status%20in%20(103,108,66,106)%20and%20statuschangetimeasstring%20>%20%27"+strTime+"%27%20and%20type%20in%20(2,6)%20and%20agentname%20like%20%27M1UTLA01%27)";
		//JobRun 
		//String apiString ="http://m1tida14.mnao.net:8080/api/tes-6.5/JobRun.getList?query=(status%20in%20(103,108,66,106)%20and%20statuschangetimeasstring%20>%20%27"+strTime+"%27%20and%20type%20in%20(2,6)%20and%20agentname%20in%20("+agent+"))";
		
		//Schedule from API
		String todaydate,proddate;
		Date date;
		Format sformatter;
		Calendar calendar = Calendar.getInstance();
		
		date = calendar.getTime();
		sformatter = new SimpleDateFormat("yyyyMMdd000000");
		todaydate = sformatter.format(date);
		System.out.println("Today's Date as String = "+todaydate);
		
		calendar.add(calendar.DATE, 1);
		date = calendar.getTime();
		sformatter = new SimpleDateFormat("yyyyMMdd000000");
		proddate = sformatter.format(date);
		System.out.println("Tomorrow's Date as String = "+proddate);
		
		String apiSchedule="http://m2tida04.mnao.net:8080/api/tes-6.5/Schedules.getList?query=(productiondateasstring="+proddate+")";
		System.out.println("API Schedule: "+apiSchedule);
		
		URL asurl = new URL(apiSchedule);
		HttpURLConnection asconn = (HttpURLConnection) asurl.openConnection();
		asconn.setRequestMethod("GET");
		asconn.setDoInput(true);
		asconn.setDoOutput(true);
		
		//Schedule from jobrun
		String apijobrun = "http://m2tida04.mnao.net:8080/api/tes-6.5/JobRun.getList?query=(productiondateasstring="+proddate+"%20and%20type%20in%20(2,6))";
		System.out.println("API Schedule: "+apijobrun);
		URL jrurl = new URL(apijobrun);
		HttpURLConnection jrconn = (HttpURLConnection) jrurl.openConnection();
		jrconn.setRequestMethod("GET");
		jrconn.setDoInput(true);
		jrconn.setDoOutput(true);
		
		//Agent from API
		 String apiString ="http://m2tida04.mnao.net:8080/api/tes-6.5/Node.getList";
		 System.out.println("API Connections: "+apiString);
		 //System.out.println("generated end point = "+apiString);
			URL url = new URL(apiString);
					
			HttpURLConnection conn = (HttpURLConnection) url.openConnection();
			conn.setRequestMethod("GET");
			conn.setDoInput(true);
			conn.setDoOutput(true);
			if (sessionID == null)
			{
				String userNamePassword = "mnao\\"+loader.getValue("CREDENTIALS");
				userNamePassword =
						new
						String(org.apache.commons.codec.binary.Base64.encodeBase64(userNamePassword
								.getBytes()));

				String basicAuth = "Basic " + userNamePassword;
				conn.setRequestProperty("Authorization", basicAuth);
				asconn.setRequestProperty("Authorization", basicAuth);
				jrconn.setRequestProperty("Authorization", basicAuth);
			}
			else
			{
				conn.setRequestProperty("Cookie", sessionID);
			}
			conn.connect();
	
		try	
		{
			
			//Creating Excel file
			
			HSSFWorkbook wb = new HSSFWorkbook();

			HSSFSheet spreadSheet = wb.createSheet("spreadSheet");

			HSSFRow row = spreadSheet.createRow(0);
			HSSFCell cell = row.createCell(0);
			cell.setCellValue("Machine");
			//row.createCell(1).setCellValue("Group Name");
			row.createCell(1).setCellValue("Agent Name");
			row.createCell(2).setCellValue("Enabled");
			row.createCell(3).setCellValue("Linked");
			row.createCell(4).setCellValue("Agent Version"); 
			
			HSSFFont font = wb.createFont();
			font.setFontHeightInPoints((short) 30);
		    font.setFontName("Calibri");

		    //Document for Test Agent Connections
			DocumentBuilderFactory factory =DocumentBuilderFactory.newInstance();
			DocumentBuilder builder = factory.newDocumentBuilder();
			Document document = builder.parse(conn.getInputStream());
			
			//System.out.println(""+document.getDocumentElement());
			
			//Parsing XML response to populate Excel

			//NodeList nodeList = document.getElementsByTagName("tes:jobrun");
			
			//Document for Schedule
			DocumentBuilderFactory sfactory =DocumentBuilderFactory.newInstance();
			DocumentBuilder sbuilder = sfactory.newDocumentBuilder();
			Document sdocument = sbuilder.parse(asconn.getInputStream());
			
			//Get all elements for schedules
			/*NodeList snodeList = sdocument.getElementsByTagName("*");
	        for (int i=0; i<snodeList.getLength(); i++) {

	            Element element = (Element) snodeList.item(i);

	            System.out.println("Found element " + element.getNodeName());

	        }*/
			
			//Data Retrive for Schedule from API
			NodeList snodeList = sdocument.getElementsByTagName("tes:id");
			int srowNum=1,adhoc=0,sjobs=0,tsjobs=0;
			String sadhoc,ssjobs;
			System.out.println("No of elements for schedule: "+snodeList.getLength());
			sadhoc = sdocument.getElementsByTagName("tes:adhocjobs2").item(0).getTextContent();
			adhoc = Integer.parseInt(sadhoc);
			System.out.println("Total no adhoc of jobs in schedule tab: " + adhoc);
			ssjobs = sdocument.getElementsByTagName("tes:scheduledjobs2").item(0).getTextContent();
			sjobs = Integer.parseInt(ssjobs);
			System.out.println("Total no of scheduled jobs in schedule tab: " + sjobs);
			tsjobs = adhoc + sjobs;
			System.out.println("Total no of all jobs in schedule tab adhoc(" + adhoc + ") + scheduled(" + sjobs + "): " + tsjobs);
			
			//Document for JobRun
			DocumentBuilderFactory jrfactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder jrbuilder = jrfactory.newDocumentBuilder();
			Document jrdocument = jrbuilder.parse(jrconn.getInputStream());
			
			//Get all elements for jobrun
			/*NodeList jrnodeList = jrdocument.getElementsByTagName("*");
	        for (int i=0; i<jrnodeList.getLength(); i++) {

	            Element jrelement = (Element) jrnodeList.item(i);

	            System.out.println("Found element " + jrelement.getNodeName());

	        }*/
			
			//Data Retirve from JobRun API
			NodeList jrnodeList = jrdocument.getElementsByTagName("tes:jobrun");
			int jrijobs;
			jrijobs = jrnodeList.getLength();
			System.out.println("No of jobs from JobRun: "+jrnodeList.getLength());
			
			//Get all elements for test agent connections
			/*NodeList nodeList = document.getElementsByTagName("*");

	         

	        for (int i=0; i<nodeList.getLength(); i++) {

	            Element element = (Element) nodeList.item(i);

	            System.out.println("Found element " + element.getNodeName());

	        }*/
			
			//Data Retrive for Test Agent Connection
			NodeList nodeList = document.getElementsByTagName("tes:node");
			int rowNum=1;
			System.out.println("No of Connections: "+nodeList.getLength());
			for(int temp= 0; temp < nodeList.getLength() ; temp ++)
			{
				Node nNode = nodeList.item(temp);

				//System.out.println(""+nodeList.item(temp));

				if(nNode.getNodeType() == Node.ELEMENT_NODE)
				{
					Element eElement = (Element) nNode;
									
					row = spreadSheet.createRow(rowNum);
					
					//System.out.println(""+eElement.getElementsByTagName("tes:machine").item(0).getTextContent());
					row.createCell(0).setCellValue(eElement.getElementsByTagName("tes:machine").item(0).getTextContent());
					row.createCell(1).setCellValue(eElement.getElementsByTagName("tes:name").item(0).getTextContent());
					row.createCell(2).setCellValue(eElement.getElementsByTagName("tes:active").item(0).getTextContent());
					try
					{
						//System.out.println(""+eElement.getElementsByTagName("tes:connectionactive").item(0).getTextContent());
						row.createCell(3).setCellValue(eElement.getElementsByTagName("tes:connectionactive").item(0).getTextContent());
					}
					catch (Exception e)
					{
						System.out.println("");
					}
					
					try
					{
						//System.out.println(""+eElement.getElementsByTagName("tes:connectionactive").item(0).getTextContent());
						row.createCell(4).setCellValue(eElement.getElementsByTagName("tes:agentversion").item(0).getTextContent());
					}
					catch (Exception e)
					{
						System.out.println("");
					}
										
					//row.createCell(0).setCellValue(eElement.getElementsByTagName("tes:machine").item(0).getTextContent());
					//if (eElement.getElementsByTagName("tes:parentname").item(0) != null)
					//{
					//row.createCell(1).setCellValue(eElement.getElementsByTagName("tes:parentname").item(0).getTextContent());
					//}
					/*row.createCell(1).setCellValue(eElement.getElementsByTagName("tes:name").item(0).getTextContent());
					row.createCell(2).setCellValue(eElement.getElementsByTagName("tes:active").item(0).getTextContent());
					row.createCell(3).setCellValue(eElement.getElementsByTagName("tes:connectionactive").item(0).getTextContent());
					row.createCell(4).setCellValue(eElement.getElementsByTagName("tes:agentversion").item(0).getTextContent()); */
				}
				rowNum ++;
			}

			spreadSheet.autoSizeColumn(0);
			spreadSheet.autoSizeColumn(1);
			spreadSheet.autoSizeColumn(2);
			spreadSheet.autoSizeColumn(3);
			spreadSheet.autoSizeColumn(4);
			spreadSheet.autoSizeColumn(5);
			
			String strCurrentDate = formatter.format(new Date());
			String filePath = loader.getValue("FILE_PATH");
			String fileName = "Tidal Test Agent Report "+ strCurrentDate+".xls";
			FileOutputStream fileOut = new FileOutputStream(filePath+File.separator+fileName);
			wb.write(fileOut);

			fileOut.close();

			wb.close();
			TestAgentScheduleReport.sendEmail(filePath, fileName,loader.getValue("RECIPIENT"),loader.getValue("RECIPIENTCC"),tsjobs,jrijobs);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}

		
	}
	
	public static void sendEmail(String filePath, String fileName, String recipient, String recipientcc, int tsjobs, int jrijobs) 
	   {    
	      
	      String sender = "noreply@tidalMazdausa.com";
	      String host = "smtp.mazdausa.com";
	 
	      Properties properties = System.getProperties();
	      properties.setProperty("mail.smtp.host", host);
	 
	      Session session = Session.getDefaultInstance(properties);
	 
	      try
	      {  
	    	  
		         MimeMessage message = new MimeMessage(session);
		         message.setFrom(new InternetAddress(sender));
		         message.addRecipients(Message.RecipientType.TO, InternetAddress.parse(recipient));
		         message.addRecipient(Message.RecipientType.CC, new InternetAddress(recipientcc)); 
		         message.setSubject("Daily Test Tidal Agent & Schedule Report");
		         String stab,activity,agentmsg,vreq;
		         stab="Hi Team, \n\nTotal no. of jobs in Schedule tab: "+tsjobs;
		         activity="\n\nTotal no. of jobs in Schedule Comaprison tab: "+jrijobs;
		         vreq="\n\nPlease validate the same.";
		         agentmsg="\n\nPlease refer to the attached report for TIDAL Test Agent Connections.\n\nThanks\nTidal Scheduling Team";
		         //String messageBody= "Hi Team, \n\nPlease refer to the attached report for TIDAL Test Agent Connections.\n\nThanks\nTidal Scheduling Team";
		         String messageBody = stab+activity+vreq+agentmsg;
		        
		         MimeBodyPart messageBodyPart = new MimeBodyPart();
		         messageBodyPart.setText(messageBody);
		         
		         Multipart multipart = new MimeMultipart();
		         
		         multipart.addBodyPart(messageBodyPart);
		         messageBodyPart = new MimeBodyPart();
		         DataSource source = new FileDataSource(filePath+File.separator+fileName);
		         messageBodyPart.setDataHandler(new DataHandler(source));
		         
		         messageBodyPart.setFileName(fileName);
		         multipart.addBodyPart(messageBodyPart);
		         message.setContent(multipart);
		         // Send email.
		         Transport.send(message);
		         System.out.println("Mail successfully sent");
	      }
	      catch (MessagingException mex) 
	      {
	         mex.printStackTrace();
	      }
	   }
	
	

}