using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.Collections;
using Microsoft.Office.DocumentManagement.DocumentSets;
using Microsoft.SharePoint.Utilities;
using System.Web;
using System.Web.UI;
using System.Collections.Specialized;

namespace PHO_Copy_Document_Set.Layouts.PHO_Copy_Document_Set
{
    public partial class Export_Document_Set : LayoutsPageBase
    {

        
        private void ActionDocLibRecursive()
        {           
           
            using (SPSite site = new SPSite(SPContext.Current.Site.Url))
            {
                using (SPWeb web = site.OpenWeb())
                { 

                    SPListCollection docLibraryColl = web.GetListsOfType(SPBaseType.DocumentLibrary);

                    // loop through each list in the web site
                    foreach (SPList list in docLibraryColl)
                    {

                        if (list.BaseType.ToString() == "DocumentLibrary" && list.BaseTemplate == SPListTemplateType.DocumentLibrary)
                        {
                            if (!list.Hidden && list.ToString() != "Style Library" && list.ToString() != "Form Templates" && list.ToString() != "Site Collection Documents")
                            {
                              DropDownList1.Items.Add(list.Title.ToString());
                            }
                        }   

                    }
                }
            }

     }

        protected void Page_Load(object sender, EventArgs e)
        {          
          ActionDocLibRecursive();
        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            Label1.Text = "";
            
            try
            {
                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                using (SPWeb web = site.OpenWeb())
                {

                    string ListId = Request.Params["ListId"];                   
                    Guid id = new Guid(ListId);
                    SPList sourceList = web.Lists.GetList(id, true);                                        
                    string ItemId = Request.Params["ItemId"]; 
                 
                    SPListItem sourceItem = sourceList.GetItemById(Convert.ToInt32(ItemId));
                    DocumentSet documentSet = DocumentSet.GetDocumentSet(sourceItem.Folder);
                    
                    SPList targetList = web.Lists[DropDownList1.SelectedItem.Text.ToString()];

                    string sourceCType = sourceItem.ContentType.Name.ToString();
                    //string sourceCTypeParentName = sourceItem.ContentType.Parent.Name;

                    SPContentTypeCollection oCTypeColl = targetList.ContentTypes;
                    StringCollection Colec = new StringCollection();
                     foreach (SPContentType conttype in oCTypeColl)
                             {                                
                              Colec.Add(conttype.Name.ToString());         
                             }
                     if (Colec.Contains(sourceCType))
                     {
                        
                         SPContentTypeId contentTypeId = targetList.ContentTypes[sourceCType].Id;
                         byte[] documentSetData = documentSet.Export();
                         string documentSetName = documentSet.Item.Name;
                         SPFolder targetFolder = targetList.RootFolder;
                         Hashtable properties = sourceItem.Properties;                        
                         DocumentSet.Import(documentSetData, documentSetName, targetFolder, contentTypeId, properties, web.CurrentUser);
                         
                         try
                         {
                             web.AllowUnsafeUpdates = true;                           
                             documentSet.VersionCollection.Add(true,"Document set item has been exported to destination library by: " + web.CurrentUser);
                             sourceItem.Update();
                             web.AllowUnsafeUpdates = false;
                         }
                         catch(Exception ex1)
                         {
                             Label1.ForeColor = System.Drawing.Color.Red;
                             Label1.Text = ex1.Message;
                         }

                         string urlRed = site.Url + "/" + targetList + "/Forms/AllItems.aspx";
                         Response.Redirect(urlRed);

                         
                     }
                     else
                     {
                         Label1.ForeColor = System.Drawing.Color.Red;
                         Label1.Text = "No content type found. Go to your destination library and add content type.";
                     }

                 }
            
            }
            catch(Exception ex)
            {
                Label1.ForeColor = System.Drawing.Color.Red;
                Label1.Text = ex.Message;

            }
        }
    }
}
