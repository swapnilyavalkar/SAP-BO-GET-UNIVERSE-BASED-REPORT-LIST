// Authored By swapnil.yavalkar //

import java.util.StringTokenizer;

import com.crystaldecisions.sdk.exception.SDKException;
import com.crystaldecisions.sdk.occa.infostore.IInfoObject;
import com.crystaldecisions.sdk.occa.infostore.IInfoObjects;
import com.crystaldecisions.sdk.occa.infostore.internal.InfoObject;
import com.crystaldecisions.sdk.plugin.desktop.folder.IFolder;
import com.crystaldecisions.sdk.properties.IProperties;
import com.crystaldecisions.sdk.properties.IProperty;


public class UniverseReports extends InitCmsExcel {

	
	public static void main(String[] args) {
		try {
			UniverseReports universe = new UniverseReports();
			
			initcmsexcel.setCMS(args[0]);
			initcmsexcel.setUserId(args[1]);
			initcmsexcel.setPassword(args[2]);
			initcmsexcel.setUniverseID(args[3]);
			
			initcmsexcel.initCMSConnection();
			initcmsexcel.initExcel();
			
			
			initcmsexcel.addRow("Report Name|Report ID|Report Type|Report Owner|Universe Name|Report Path");
			
			universe.getReporIds(initcmsexcel.universeid);
			initcmsexcel.setFilename();
			initcmsexcel.saveExcel();
			
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private void getReporIds(String universeid) throws SDKException {
		
		String temp="";
		String unvname;
		int index;
		
		String sq = "SELECT TOP 50000 SI_NAME, SI_SL_DOCUMENTS,SI_WEBI from CI_INFOOBJECTS, CI_SYSTEMOBJECTS, CI_APPOBJECTS where SI_ID =" + universeid;
		IInfoObjects iObjects = iStore.query(sq);
		InfoObject iobjects;
		for (int i = 0; i < iObjects.size(); i++)
		{
			
			iobjects = (InfoObject) iObjects.get(i);
			
			IProperties props=iobjects.properties();
			IProperty reportID=(IProperty) props.getProperty("SI_SL_DOCUMENTS");
		
			if (reportID == null)
			{
				props=iobjects.properties();
				reportID=(IProperty) props.getProperty("SI_WEBI");
				
				StringTokenizer eq = new StringTokenizer(reportID.toString(), ",", false);
				
				while(eq.hasMoreTokens()){
					temp = eq.nextToken();
					index = temp.indexOf("=");	
					temp = temp.substring((index+1),temp.length());
					
					
					IProperty universename = (IProperty) props.getProperty("SI_NAME");
					unvname = universename.toString();
					
					getReportDetails(temp,unvname);
					
					}	
			
			}else
			{
				StringTokenizer eq = new StringTokenizer(reportID.toString(), ",", false);
				
				while(eq.hasMoreTokens()){
					temp = eq.nextToken();
					index = temp.indexOf("=");	
					temp = temp.substring((index+1),temp.length());
					
					
					IProperty universename = (IProperty) props.getProperty("SI_NAME");
					unvname = universename.toString();
					
					getReportDetails(temp,unvname);
					
					}	
			}
			
		}

 }
		

	private void getReportDetails(String reportid, String unvname) throws SDKException {
		
		String sq = "SELECT TOP 50000 SI_NAME, SI_ID, SI_KIND, SI_OWNER, SI_PARENTID, SI_FILES, SI_PARENT_FOLDER from CI_INFOOBJECTS, CI_SYSTEMOBJECTS, CI_APPOBJECTS WHERE si_kind IN ('FullClient', 'Webi', 'CrystalReport') AND SI_INSTANCE = 0 AND SI_ID =" + reportid;
		
		initcmsexcel.setUniverseName(unvname);
		
		IInfoObjects iObjects = iStore.query(sq);
		IInfoObject iObject = null;
		
		for (int i = 0; i < iObjects.size(); i++)
		{
			StringBuilder rowData = new StringBuilder();
			
			iObject = (IInfoObject) iObjects.get(i);
			
			
			rowData.append(iObject.getTitle());
			rowData.append("|");
			
			rowData.append(iObject.getID());
			rowData.append("|");
			
			rowData.append(iObject.getKind());
			rowData.append("|");
			
			rowData.append(iObject.getOwner());
			rowData.append("|");
			
			rowData.append(unvname);
			rowData.append("|");
						
			IProperties prop = iObject.properties();
			IProperty getProp = prop.getProperty("SI_PARENTID");
			String FolderID = getProp.toString();
			IInfoObjects folder = iStore.query("select si_id,si_CUID,si_name,si_parentid,si_path from ci_infoobjects where si_id="+ FolderID);
			IInfoObject ifolder = (IInfoObject) folder.get(0);
			if (ifolder.getKind().equals("Folder")) {
				IFolder iifolder = (IFolder) ifolder;
				String finalFolderPath = "";
				if (iifolder.getPath() != null) {
					String path[] = iifolder.getPath();
					for (int fi = 0; fi < path.length; fi++) {
						finalFolderPath = path[fi] + "/" + finalFolderPath;
					}
					finalFolderPath = finalFolderPath + iifolder.getTitle();
				} else {
					finalFolderPath = finalFolderPath + iifolder.getTitle();
				}

				rowData.append(finalFolderPath);
			} else if ((ifolder.getKind().equals("FavoritesFolder"))) {
				rowData.append("FavoritesFolder ::  ");
				rowData.append(ifolder.getTitle());
			} else if ((ifolder.getKind().equals("Inbox"))) {
				rowData.append(" Inbox ::  ");
				rowData.append(ifolder.getTitle());
			} else if ((ifolder.getKind().equals("ObjectPackage"))) {
				IProperties prop1 = ifolder.properties();
				IProperty getProp1 = prop1.getProperty("SI_PARENTID");
				String FolderID1 = getProp1.toString();
				IInfoObjects folder1 = iStore
						.query("select * from ci_infoobjects where si_id="
								+ FolderID1);
				IInfoObject ifolder1 = (IInfoObject) folder1.get(0);
				if (ifolder1.getKind().equals("Folder")) {
					IFolder iifolder1 = (IFolder) ifolder1;
					String finalFolderPath1 = "";
					if (iifolder1.getPath() != null) {
						String path[] = iifolder1.getPath();
						for (int j = 0; j < path.length; j++) {
							finalFolderPath1 = path[j] + "/" + finalFolderPath1;
						}
						finalFolderPath1 = finalFolderPath1
								+ iifolder1.getTitle() + "/"
								+ ifolder.getTitle();
					} else {
						finalFolderPath1 = finalFolderPath1
								+ iifolder1.getTitle() + "/"
								+ ifolder.getTitle();
					}
					rowData.append(finalFolderPath1);
				}
			}
			initcmsexcel.addRow(rowData.toString());
			
		}	
}
		
}
