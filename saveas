Outlook.Inspector inspector = null; 
object item = null; 
try { 
    inspector = OutlookApp.ActiveInspector(); 
    if (inspector != null) { 
    item = inspector.CurrentItem;  
    if (item is Outlook.MailItem) { 
        Outlook.MailItem mail = item as Outlook.MailItem; 
        // do your stuff here 
        string fileName = string.Format(@"D:{0}.msg", Guid.NewGuid().ToString("N"));  
        mailItem.SaveAs(fileName); 
    } 
} catch (Exception ex) { 
    // log exception here 
} finally {  
    if (item != null) Marshal.ReleaseComObject(item); // in .NET, item and mail refer to the same COM object  
    if (inspector != null) Marshal.ReleaseComObject(inspector);  
} 
