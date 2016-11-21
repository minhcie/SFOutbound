package src.main.java;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.UUID;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.partner.Connector;
import com.sforce.soap.partner.DeleteResult;
import com.sforce.soap.partner.PartnerConnection;
import com.sforce.soap.partner.QueryResult;
import com.sforce.soap.partner.SaveResult;
import com.sforce.soap.partner.Error;
import com.sforce.soap.partner.sobject.SObject;
import com.sforce.ws.ConnectionException;
import com.sforce.ws.ConnectorConfig;

import org.apache.log4j.Logger;

public class SFOutbound {
    private static final Logger log = Logger.getLogger(SFOutbound.class.getName());
    private static final String USERNAME = "mtran@211sandiego.org";
    private static final String PASSWORD = "m1nh@211KsmlvVA4mvtI6YwzKZOLjbKF9";
    private static PartnerConnection connection;

    static void usage() {
        System.err.println("");
        System.err.println("usage: java -jar SFOutbound.jar <data.xlsx>");
        System.err.println("");
        System.exit(-1);
    }

    public static void main(String[] args) {
        if (args.length == 0 || args.length < 1) {
            usage();
        }

    	ConnectorConfig config = new ConnectorConfig();
    	config.setUsername(USERNAME);
    	config.setPassword(PASSWORD);
    	//config.setTraceMessage(true);

        try {
            // Check to make sure excel data file exists.
            String fileName = args[0];
            log.info("Reading excel data file " + fileName + "...");
            File xcel = new File(fileName);
            if (!xcel.exists()) {
                log.error("Data file " + fileName + " not found!");
                System.exit(-1);
            }

            // Read data file.
            List<ContactInfo> data = readData(xcel);
            if (data == null || data.size() <= 0) {
                log.info("No data found!");
                System.exit(0);
            }

            // Establish Salesforce web service connection.
    		connection = Connector.newConnection(config);

    		// @debug.
    		log.info("Auth EndPoint: " + config.getAuthEndpoint());
    		log.info("Service EndPoint: " + config.getServiceEndpoint());
    		log.info("Username: " + config.getUsername());
    		log.info("SessionId: " + config.getSessionId());

            // Query default owner.
            String owner = "Karla Samayoa";
            String ownerId = queryUser(connection, owner);
            if (ownerId == null) {
                log.error("Invalid owner name!");
                System.exit(-1);
            }

            // Query record type id.
            String acctId = queryAccount(connection, "ALL CLIENTS");
            String contactRecordTypeId = queryRecordType(connection, "Contact",
                                                         "Client");
            String campaignRecordTypeId = queryRecordType(connection, "Campaign",
                                                          "Call Center");
            String campaignMemberRecordTypeId = queryRecordType(connection, "CampaignMember",
                                                                "Client");

            // Insert/update campaign.
            Date today = new Date();
            SimpleDateFormat sdf = new SimpleDateFormat("MMyyyy");
            ContactInfo ci = data.get(0);
            String campaignName = "Campaign_SF_" + sdf.format(today);
            String campaignId = queryCampaign(connection, campaignName);
            if (campaignId == null) {
                campaignId = createCampaign(connection, campaignRecordTypeId,
                                            campaignMemberRecordTypeId, ownerId,
                                            campaignName, ci.extractDate);
            }

            // Insert/update contacts.
            for (int i = 0; i < data.size(); i++) {
                ci = data.get(i);

                String contactId = null;
                SObject so = queryContact(connection, acctId, ci.firstName,
                                          ci.lastName, ci.caseId);
                if (so != null) {
                    contactId = so.getId();
                    updateContact(connection, contactId, ci);
                }
                else {
                    contactId = createContact(connection, acctId, contactRecordTypeId,
                                              ownerId, ci);
                }

                // Insert campaign member.
                so = queryCampaignMember(connection, campaignId, contactId);
                if (so == null) {
                    createCampaignMember(connection, campaignId, contactId);
                }
            }
        }
    	catch (ConnectionException e) {
            log.error(e.getMessage());
            e.printStackTrace();
    	}
        catch (IOException ioe) {
            log.error(ioe.getMessage());
            ioe.printStackTrace();
        }
        catch (Exception e) {
            log.error(e.getMessage());
            e.printStackTrace();
        }
    }

    private static String isNumberOrDate(Cell cell) {
        String retVal;
        if (HSSFDateUtil.isCellDateFormatted(cell)) {
            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
            retVal = sdf.format(cell.getDateCellValue());
        }
        else {
            DataFormatter formatter = new DataFormatter();
            retVal = formatter.formatCellValue(cell);
        }
        return retVal;
    }

    private static String getCellValue(Cell cell) {
        String retVal = "";
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                retVal = "" + cell.getBooleanCellValue();
                break;

            case Cell.CELL_TYPE_STRING:
                retVal = cell.getStringCellValue();
                break;

            case Cell.CELL_TYPE_NUMERIC:
                retVal = isNumberOrDate(cell);
                break;

            case Cell.CELL_TYPE_BLANK:
            default:
                retVal = "";
        }
        return retVal.trim();
    }

    private static List<ContactInfo> readData(File xcel) throws Exception {
        List<ContactInfo> data = new ArrayList<ContactInfo>();

        // Get the workbook object for xlsx file.
        XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(xcel));

        // Get correct sheet from the workbook.
        XSSFSheet sheet = wBook.getSheetAt(0); // The only sheet.
        Iterator<Row> rowIterator = sheet.iterator();

        SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
        Row row;
        Cell cell;

        while (rowIterator.hasNext()) {
            row = rowIterator.next();

            // Ignore column header row.
            if (row.getRowNum() == 0) {
                continue;
            }

            // @debug.
            /*
            if (row.getRowNum() >= 4) {
                break;
            }
            */

            // Row data.
            boolean hasData = false;
            ContactInfo ci = new ContactInfo();

            // Iterate to all cells (including empty cell).
            int minColIndex = row.getFirstCellNum();
            int maxColIndex = row.getLastCellNum();
            for (int colIndex = minColIndex; colIndex < maxColIndex; colIndex++) {
                String cellValue = "";
                cell = row.getCell(colIndex);
                if (cell == null) {
                    log.info("row " + row.getRowNum() + " col " + colIndex + " is null");
                }
                else {
                    cellValue = getCellValue(cell);
                    if (cellValue.equalsIgnoreCase("null")) {
                        cellValue = "";
                    }
                }

                if (cellValue.length() > 0) {
                    hasData = true;
                }

                switch (colIndex) {
                    case 0:
                        if (cellValue.length() > 0) {
                            ci.extractDate = sdf.parse(cellValue);
                        }
                        break;
                    case 1:
                        ci.caseId = cellValue;
                        break;
                    case 2:
                        ci.language = cellValue;
                        break;
                    case 3:
                        ci.firstName = cellValue;
                        break;
                    case 4:
                        ci.lastName = cellValue;
                        break;
                    case 8:
                        ci.address = cellValue;
                        break;
                    case 9:
                        ci.city = cellValue;
                        break;
                    case 10:
                        ci.state = cellValue;
                        break;
                    case 11:
                        ci.zip = cellValue;
                        break;
                    case 12:
                        ci.phone1 = cellValue;
                        break;
                    case 13:
                        ci.phone2 = cellValue;
                        break;
                    case 16:
                        ci.email = cellValue;
                        break;
                    default:
                        break;
                }

                // Done with this row.
                if (colIndex >= 16) {
                    break;
                }
            }

            // Add only if row has data.
            if (hasData) {
                data.add(ci);
            }
        }
        return data;
    }

    private static String queryRecordType(PartnerConnection conn, String objectName,
                                          String name) {
    	log.info("Querying for " + name + " record type from " + objectName + "...");
        String recordTypeId = null;
    	try {
            // Query for record type name.
    		String sql = "SELECT Id, Name, SobjectType FROM RecordType " +
                         "WHERE Name = '" + name + "' " +
                         "  AND SobjectType = '" + objectName + "' ";
    		QueryResult queryResults = conn.query(sql);
    		if (queryResults.getSize() > 0) {
    			for (SObject s: queryResults.getRecords()) {
                    recordTypeId = s.getId();
    			}
    		}
    	}
    	catch (Exception e) {
    		e.printStackTrace();
    	}

        // @debug.
        if (recordTypeId != null) {
			log.info("Record Type Id: " + recordTypeId);
        }
        else {
            log.info(name + " record type not found!");
        }
        return recordTypeId;
    }

    private static String queryUser(PartnerConnection conn, String name) {
        log.info("Querying user " + name + "...");
        String userId = null;
        if (name == null || name.trim().length() <= 0) {
            return userId;
        }

        // Parse user first and last name.
        String[] parts = name.split(" ");

    	try {
    		StringBuilder sb = new StringBuilder();
    		sb.append("SELECT Id, FirstName, LastName, Account.Name ");
    		sb.append("FROM User ");
    		sb.append("WHERE FirstName = '" + parts[0] + "' ");
    		sb.append("  AND LastName = '" + parts[1] + "'");

    		QueryResult queryResults = conn.query(sb.toString());
    		if (queryResults.getSize() > 0) {
    			for (SObject s: queryResults.getRecords()) {
                    userId = s.getId();
    			}
    		}
    	}
    	catch (Exception e) {
    		e.printStackTrace();
    	}
        return userId;
    }

    private static String queryAccount(PartnerConnection conn, String name) {
    	log.info("Querying for account name " + name);
        String acctId = null;
    	try {
            // Query for record type name.
    		String sql = "SELECT Id, Name FROM Account " +
                         "WHERE Name = '" + name + "' ";
    		QueryResult queryResults = conn.query(sql);
    		if (queryResults.getSize() > 0) {
    			for (SObject s: queryResults.getRecords()) {
                    acctId = s.getId();
    			}
    		}
    	}
    	catch (Exception e) {
    		e.printStackTrace();
    	}

        // @debug.
        if (acctId != null) {
			log.info("Account Id: " + acctId);
        }
        else {
            log.info(name + " account not found!");
        }
        return acctId;
    }

    private static String queryCampaign(PartnerConnection conn, String name) {
    	log.info("Querying campaign name " + name + "...");
        String campaignId = null;
    	try {
    		StringBuilder sb = new StringBuilder();
    		sb.append("SELECT Id, Name ");
    		sb.append("FROM Campaign ");
    		sb.append("WHERE IsActive = TRUE ");
    		sb.append("  AND Name = '" + name + "' ");
    		QueryResult queryResults = connection.query(sb.toString());
    		if (queryResults.getSize() > 0) {
    			for (SObject s: queryResults.getRecords()) {
                    campaignId = s.getId();
    			}
    		}
    	}
    	catch (Exception e) {
    		e.printStackTrace();
    	}
        return campaignId;
    }

    private static SObject queryCampaignMember(PartnerConnection conn,
                                               String campaignId, String contactId) {
    	log.info("Querying campaign member...");
        SObject result = null;
        try {
    		StringBuilder sb = new StringBuilder();
    		sb.append("SELECT Id, CampaignId, ContactId, Name ");
    		sb.append("FROM CampaignMember ");
    		sb.append("WHERE CampaignId = '" + campaignId + "' ");
    		sb.append("  AND ContactId = '" + contactId + "'");

    		QueryResult queryResults = conn.query(sb.toString());
    		if (queryResults.getSize() > 0) {
                result = queryResults.getRecords()[0];
    		}
    	}
    	catch (Exception e) {
    		e.printStackTrace();
    	}
        return result;
    }

    private static SObject queryContact(PartnerConnection conn, String acctId,
                                        String firstName, String lastName, String caseId) {
    	log.info("Querying contact " + firstName + " " + lastName + "...");
        SObject result = null;
        try {
    		StringBuilder sb = new StringBuilder();
    		sb.append("SELECT Id, FirstName, LastName, AccountId, Mailing_City__c ");
    		sb.append("FROM Contact ");
    		sb.append("WHERE AccountId = '" + acctId + "' ");
    		sb.append("  AND FirstName = '" + firstName + "' ");
    		sb.append("  AND LastName = '" + lastName + "' ");
    		sb.append("  AND County_Case_ID__c = '" + caseId + "'");

    		QueryResult queryResults = conn.query(sb.toString());
    		if (queryResults.getSize() > 0) {
                result = queryResults.getRecords()[0];
    		}
    	}
    	catch (Exception e) {
    		e.printStackTrace();
    	}
        return result;
    }

    private static String createCampaign(PartnerConnection conn, String campaignRecordTypeId,
                                         String campaignMemberRecordTypeId, String ownerId,
                                         String campaignName, Date extractDate) {
        log.info("Creating new campaign name: " + campaignName);
        String campaignId = null;
    	try {
    	    SObject[] records = new SObject[1];

            SObject so = new SObject();
    		so.setType("Campaign");
	        so.setField("RecordTypeId", campaignRecordTypeId);;
	        so.setField("CampaignMemberRecordTypeId", campaignMemberRecordTypeId);;
	        so.setField("OwnerId", ownerId);;
    		so.setField("Name", campaignName);
    		so.setField("IsActive", new Boolean(true));
    		so.setField("Type", "SF Outbound");
            so.setField("Extracted_Date__c", extractDate);
    		records[0] = so;

    		// Create the records in Salesforce.
    		SaveResult[] saveResults = conn.create(records);

    		// Check the returned results for any errors.
    		for (int i = 0; i < saveResults.length; i++) {
    			if (saveResults[i].isSuccess()) {
    				campaignId = saveResults[i].getId();
    				log.info(i + ". Successfully created record - Id: " + campaignId);
    			}
    			else {
    				Error[] errors = saveResults[i].getErrors();
    				for (int j = 0; j< errors.length; j++) {
    					log.error(errors[j].getStatusCode() + " - " + errors[j].getMessage());
    				}
    			}
    		}
    	}
    	catch (Exception e) {
    		e.printStackTrace();
    	}
        return campaignId;
    }

    private static String createCampaignMember(PartnerConnection conn, String campaignId,
                                               String contactId) {
        log.info("Adding new campaign member id: " + contactId);
        String campaignMemberId = null;
    	try {
    	    SObject[] records = new SObject[1];

            SObject so = new SObject();
    		so.setType("CampaignMember");
    		so.setField("CampaignId", campaignId);
    		so.setField("ContactId", contactId);
    		records[0] = so;

    		// Create the records in Salesforce.
    		SaveResult[] saveResults = conn.create(records);

    		// Check the returned results for any errors.
    		for (int i = 0; i < saveResults.length; i++) {
    			if (saveResults[i].isSuccess()) {
    				campaignMemberId = saveResults[i].getId();
    				log.info(i + ". Successfully created record - Id: " + campaignMemberId);
    			}
    			else {
    				Error[] errors = saveResults[i].getErrors();
    				for (int j = 0; j< errors.length; j++) {
    					log.error(errors[j].getStatusCode() + " - " + errors[j].getMessage());
    				}
    			}
    		}
    	}
    	catch (Exception e) {
    		e.printStackTrace();
    	}
        return campaignMemberId;
    }

    private static String createContact(PartnerConnection conn, String acctId,
                                        String contactRecordTypeId, String ownerId,
                                        ContactInfo ci) {
        log.info("Creating new contact name: " + ci.firstName + " " + ci.lastName);
        String contactId = null;
    	try {
            SObject[] records = new SObject[1];

			SObject so = copyContactInfo(ci);
	        so.setField("AccountId", acctId);
	        so.setField("RecordTypeId", contactRecordTypeId);;
	        so.setField("OwnerId", ownerId);;
            records[0] = so;

            // Create the record in Salesforce.
            SaveResult[] saveResults = connection.create(records);

    		// Check the returned results for any errors.
    		for (int i = 0; i < saveResults.length; i++) {
    			if (saveResults[i].isSuccess()) {
    				contactId = saveResults[i].getId();
    				log.info(i + ". Successfully created record - Id: " + contactId);
    			}
    			else {
    				Error[] errors = saveResults[i].getErrors();
    				for (int j = 0; j< errors.length; j++) {
    					log.error(errors[j].getStatusCode() + " - " + errors[j].getMessage());
    				}
    			}
    		}
    	}
    	catch (Exception e) {
    		e.printStackTrace();
    	}
        return contactId;
    }

    private static void updateContact(PartnerConnection conn, String contactId,
                                      ContactInfo ci) {
    	log.info("Updating contact Id: " + contactId + "...");
    	SObject[] records = new SObject[1];
    	try {
			SObject so = copyContactInfo(ci);
			so.setId(contactId);
			records[0] = so;

    		// Update the record in Salesforce.
    		SaveResult[] saveResults = conn.update(records);

    		// Check the returned results for any errors.
    		for (int i = 0; i < saveResults.length; i++) {
    			if (saveResults[i].isSuccess()) {
    				log.info("Successfully updated record - Id: " + saveResults[i].getId() + "\n");
    			}
    			else {
    				Error[] errors = saveResults[i].getErrors();
    				for (int j = 0; j < errors.length; j++) {
    					log.error(errors[j].getStatusCode() + " - " + errors[j].getMessage());
    				}
    			}
    		}
    	}
    	catch (Exception e) {
    		e.printStackTrace();
    	}
    }

    private static SObject copyContactInfo(ContactInfo ci) {
        SObject so = new SObject();
		so.setType("Contact");
        so.setField("FirstName", ci.firstName);
        so.setField("LastName", ci.lastName);
		so.setField("Mailing_Street__c", ci.address);
		so.setField("Mailing_City__c", ci.city);
		so.setField("Mailing_State__c", ci.state);
		so.setField("Mailing_Zip__c", ci.zip);
        if (ci.phone1 != null) {
            so.setField("Phone_1_Primary__c", ci.phone1);
        }
        if (ci.phone2 != null) {
            so.setField("Phone_2__c", ci.phone2);
        }
        so.setField("Email", ci.email);
	    so.setField("What_is_your_preferred_language__c", ci.language);
	    //so.setField("What_Languages_do_you_Speak__c", ci.language);
	    so.setField("County_Case_ID__c", ci.caseId);
        return so;
    }
}
