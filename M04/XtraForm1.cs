using System;
using System.Text;
using DBConnection;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using MDS00;
using System.Drawing;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors;

namespace M04
{
    public partial class XtraForm1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public XtraForm1()
        {
            InitializeComponent();
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
            iniConfig = new IniFile("Config.ini");
            UserLookAndFeel.Default.SetSkinStyle(iniConfig.Read("SkinName", "DevExpress"), iniConfig.Read("SkinPalette", "DevExpress"));
        }

        private IniFile iniConfig;

        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            LookAndFeelChangedEventArgs lookAndFeelChangedEventArgs = (DevExpress.LookAndFeel.LookAndFeelChangedEventArgs)e;
            //MessageBox.Show("MyStyleChanged: " + lookAndFeelChangedEventArgs.Reason.ToString() + ", " + userLookAndFeel.SkinName + ", " + userLookAndFeel.ActiveSvgPaletteName);
            iniConfig.Write("SkinName", userLookAndFeel.SkinName, "DevExpress");
            iniConfig.Write("SkinPalette", userLookAndFeel.ActiveSvgPaletteName, "DevExpress");
        }

        private void XtraForm1_Load(object sender, EventArgs e)
        {
            bbiNew.PerformClick();
        }

        private void NewData()
        {
            txeID.EditValue = new DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCUST), '') = '' THEN 1 ELSE MAX(OIDCUST) + 1 END AS NewNo FROM Customer").getString();
            lblStatus.Text = "* New Customer";
            lblStatus.ForeColor = Color.Green;
            glueCode.EditValue = "";
            txeName.EditValue = "";
            txeShortName.EditValue = "";
            txeContacts.EditValue = "";
            txeEmail.EditValue = "";
            txeAddr1.EditValue = "";
            txeAddr2.EditValue = "";
            txeAddr3.EditValue = "";
            txeCountry.EditValue = "";
            txeTelNo.EditValue = "";
            txeFaxNo.EditValue = "";
            glueCustType.EditValue = "";
            glueSection.EditValue = "";
            glueTerm.EditValue = "";
            glueCurrency.EditValue = "";
            glueCalendar.EditValue = "";
            txeEval.EditValue = "";
            txeOthContract.EditValue = "";
            txeOthAddr1.EditValue = "";
            txeOthAddr2.EditValue = "";
            txeOthAddr3.EditValue = "";
            txeCREATE.EditValue = "0";
            txeCDATE.EditValue = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            txeUPDATE.EditValue = "0";
            txeUDATE.EditValue = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            glueCode.Focus();
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code, Name, ShortName ");
            sbSQL.Append("FROM  Customer ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT N'' AS Code, N'' AS Name, N'' AS ShortName ");
            sbSQL.Append("ORDER BY Code, Name ");
            new ObjDevEx.setGridLookUpEdit(glueCode, sbSQL, "Code", "Code").getData(true);

            //Customer Type
            sbSQL.Clear();
            sbSQL.Append("SELECT '0' AS ID, 'Customer' AS Type ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '1' AS ID, 'Sub Contract' AS Type ");
            new ObjDevEx.setGridLookUpEdit(glueCustType, sbSQL, "Type", "ID").getData(true);

            //Sales Section
            sbSQL.Clear();
            sbSQL.Append("SELECT OIDDepartment AS ID, Department ");
            sbSQL.Append("FROM Department ");
            sbSQL.Append("ORDER BY OIDDepartment ");
            new ObjDevEx.setGridLookUpEdit(glueSection, sbSQL, "Department", "Department").getData(true);

            //Payment Term
            sbSQL.Clear();
            sbSQL.Append("SELECT Name, Description ");
            sbSQL.Append("FROM PaymentTerm ");
            sbSQL.Append("ORDER BY OIDPayment ");
            new ObjDevEx.setGridLookUpEdit(glueTerm, sbSQL, "Name", "Name").getData(true);

            //Payment Currency
            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCURR AS ID, Currency ");
            sbSQL.Append("FROM Currency ");
            sbSQL.Append("ORDER BY OIDCURR ");
            new ObjDevEx.setGridLookUpEdit(glueCurrency, sbSQL, "Currency", "Currency").getData(true);

            //Calendar No.
            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCALENDAR AS No, 'Thai Parfun' AS CompanyType, 'Thai Parfun' AS CompanyName, Year  ");
            sbSQL.Append("FROM CalendarMaster  ");
            sbSQL.Append("WHERE CompanyType = 0  ");
            sbSQL.Append("UNION ALL  ");
            sbSQL.Append("SELECT OIDCALENDAR AS No, 'Customer' AS CompanyType, CompanyName, Year  ");
            sbSQL.Append("FROM CalendarMaster A  ");
            sbSQL.Append("CROSS APPLY(SELECT ShortName AS CompanyName FROM Customer WHERE OIDCUST = A.OIDCompany) B  ");
            sbSQL.Append("WHERE CompanyType = 1  ");
            sbSQL.Append("UNION ALL  ");
            sbSQL.Append("SELECT OIDCALENDAR AS No, 'Vendor' AS CompanyType, CompanyName, Year  ");
            sbSQL.Append("FROM CalendarMaster C  ");
            sbSQL.Append("CROSS APPLY(SELECT Name AS CompanyName FROM Vendor WHERE OIDVEND = C.OIDCompany) D  ");
            sbSQL.Append("WHERE CompanyType = 2  ");
            sbSQL.Append("ORDER BY Year DESC, CompanyType, CompanyName, OIDCALENDAR  ");
            new ObjDevEx.setGridLookUpEdit(glueCalendar, sbSQL, "CompanyName", "No").getData(true);

            //All Customer
            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCUST AS No, Code AS Customer, Name AS CustomerName, ShortName, Contacts AS ContactName, Email, Address1, Address2, Address3, Country, TelephoneNo, ");
            sbSQL.Append("       FaxNo, CustomerType, SalesSection, PaymentTerm, PaymentCurrency, CalendarNo, EvalutionPoint AS CustomerEvalutionPoint, OtherContact AS OtherContactName, ");
            sbSQL.Append("       OtherAddress1, OtherAddress2, OtherAddress3, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate ");
            sbSQL.Append("FROM Customer ");
            sbSQL.Append("ORDER BY CustomerType, Code ");
            new ObjDevEx.setGridControl(gcCustomer, gvCustomer, sbSQL).getData(false, false, true, true);



        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            NewData();
            LoadData();
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (glueCode.EditValue.ToString() == "")
            {
                FUNC.msgWarning("Please input customer code.");
                glueCode.Focus();
            }
            else if (txeName.EditValue.ToString() == "")
            {
                FUNC.msgWarning("Please input customer name.");
                txeName.Focus();
            }
            else if (txeShortName.EditValue.ToString() == "")
            {
                FUNC.msgWarning("Please input short name.");
                txeShortName.Focus();
            }
            else if (glueCustType.EditValue.ToString() == "")
            {
                FUNC.msgWarning("Please select customer type.");
                glueCustType.Focus();
            }
            else if (glueCalendar.EditValue.ToString() == "")
            {
                FUNC.msgWarning("Please select calendar no.");
                glueCalendar.Focus();
            }
            else
            {
                if (FUNC.msgQuiz("Confirm save data ?") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();

                    string strCREATE = "0";
                    if (txeCREATE.EditValue != null)
                    {
                        strCREATE = txeCREATE.EditValue.ToString();
                    }

                    string strUPDATE = "0";
                    if (txeUPDATE.EditValue != null)
                    {
                        strUPDATE = txeUPDATE.EditValue.ToString();
                    }

                    sbSQL.Append("IF NOT EXISTS(SELECT Code FROM Customer WHERE Code = N'" + glueCode.Text.Trim() + "') ");
                    sbSQL.Append(" BEGIN ");
                    sbSQL.Append("  INSERT INTO Customer(Code, Name, ShortName, Contacts, Email, Address1, Address2, Address3, Country, PostCode, TelephoneNo, FaxNo, CustomerType, SalesSection, PaymentTerm, PaymentCurrency, CalendarNo, EvalutionPoint, OtherContact, OtherAddress1, OtherAddress2, OtherAddress3, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) ");
                    sbSQL.Append("  VALUES(N'" + glueCode.Text.Trim() + "', N'" + txeName.Text.Trim() + "', N'" + txeShortName.Text.Trim() + "', N'" + txeContacts.Text.Trim() + "', N'" + txeEmail.Text.Trim() + "', N'" + txeAddr1.Text.Trim() + "', N'" + txeAddr2.Text.Trim() + "', N'" + txeAddr3.Text.Trim() + "', N'" + txeCountry.Text.Trim() + "', N'', N'" + txeTelNo.Text.Trim() + "', ");
                    sbSQL.Append("         N'" + txeFaxNo.Text.Trim() + "', '" + glueCustType.EditValue.ToString() + "', N'" + glueSection.Text.Trim() + "', N'" + glueTerm.Text.Trim() + "', N'" + glueCurrency.Text.Trim() + "', '" + glueCalendar.EditValue.ToString() + "', N'" + txeEval.Text.Trim() + "', N'" + txeOthContract.Text.Trim() + "', N'" + txeOthAddr1.Text.Trim() + "', N'" + txeOthAddr2.Text.Trim() + "', N'" + txeOthAddr3.Text.Trim() + "', '" + strCREATE + "', GETDATE(), '" + strUPDATE + "', GETDATE()) ");
                    sbSQL.Append(" END ");
                    sbSQL.Append("ELSE ");
                    sbSQL.Append(" BEGIN ");
                    sbSQL.Append("  UPDATE Customer SET ");
                    sbSQL.Append("      Code = N'" + glueCode.Text.Trim() + "', Name = N'" + txeName.Text.Trim() + "', ShortName = N'" + txeShortName.Text.Trim() + "', Contacts = N'" + txeContacts.Text.Trim() + "', Email = N'" + txeEmail.Text.Trim() + "', Address1 = N'" + txeAddr1.Text.Trim() + "', Address2 = N'" + txeAddr2.Text.Trim() + "', Address3 = N'" + txeAddr3.Text.Trim() + "', ");
                    sbSQL.Append("      Country = N'" + txeCountry.Text.Trim() + "', PostCode = N'', TelephoneNo = N'" + txeTelNo.Text.Trim() + "', FaxNo = N'" + txeFaxNo.Text.Trim() + "', CustomerType = '" + glueCustType.EditValue.ToString() + "', SalesSection = N'" + glueSection.Text.Trim() + "', PaymentTerm = N'" + glueTerm.Text.Trim() + "', ");
                    sbSQL.Append("      PaymentCurrency = N'" + glueCurrency.Text.Trim() + "', CalendarNo = '" + glueCalendar.EditValue.ToString() + "', EvalutionPoint = N'" + txeEval.Text.Trim() + "', OtherContact = N'" + txeOthContract.Text.Trim() + "', OtherAddress1 = N'" + txeOthAddr1.Text.Trim() + "', OtherAddress2 = N'" + txeOthAddr2.Text.Trim() + "', OtherAddress3 = N'" + txeOthAddr3.Text.Trim() + "', ");
                    sbSQL.Append("      UpdatedBy = '" + strUPDATE +"', UpdatedDate = GETDATE() ");
                    sbSQL.Append("  WHERE(OIDCUST = '" + txeID.Text.Trim() + "') ");
                    sbSQL.Append(" END ");
                    //MessageBox.Show(sbSQL.ToString());
                    if (sbSQL.Length > 0)
                    {
                        try
                        {
                            bool chkSAVE = new DBQuery(sbSQL).runSQL();
                            if (chkSAVE == true)
                            {
                                FUNC.msgInfo("Save complete.");
                                bbiNew.PerformClick();
                            }
                        }
                        catch (Exception)
                        { }
                    }
                }

            }
        }

        private void glueCode_EditValueChanged(object sender, EventArgs e)
        {
            txeName.Focus();
        }

        private void glueCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeName.Focus();
            }
        }

        private void glueCode_LostFocus(object sender, EventArgs e)
        {
            string strCODE = glueCode.Text.ToUpper().Trim();
            txeID.EditValue = "";
            lblStatus.Text = "* New Customer";
            lblStatus.ForeColor = Color.Green;
            txeName.EditValue = "";
            txeShortName.EditValue = "";
            txeContacts.EditValue = "";
            txeEmail.EditValue = "";
            txeAddr1.EditValue = "";
            txeAddr2.EditValue = "";
            txeAddr3.EditValue = "";
            txeCountry.EditValue = "";
            txeTelNo.EditValue = "";
            txeFaxNo.EditValue = "";
            glueCustType.EditValue = "";
            glueSection.EditValue = "";
            glueTerm.EditValue = "";
            glueCurrency.EditValue = "";
            glueCalendar.EditValue = "";
            txeEval.EditValue = "";
            txeOthContract.EditValue = "";
            txeOthAddr1.EditValue = "";
            txeOthAddr2.EditValue = "";
            txeOthAddr3.EditValue = "";
            txeCREATE.EditValue = "0";
            txeCDATE.EditValue = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            txeUPDATE.EditValue = "0";
            txeUDATE.EditValue = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            string gCode = glueCode.Text.ToUpper().Trim();

            if (glueCode.Text != "")
            {
                StringBuilder sbSQLx = new StringBuilder();
                sbSQLx.Append("SELECT OIDCUST FROM Customer WHERE (Code=N'" + gCode + "') ");
                string chkCode = new DBQuery(sbSQLx).getString();

                if (chkCode == "")
                {
                    sbSQLx.Clear();
                    sbSQLx.Append("SELECT Code, Name, ShortName ");
                    sbSQLx.Append("FROM  Customer ");
                    sbSQLx.Append("UNION ALL ");
                    sbSQLx.Append("SELECT N'' AS Code, N'' AS Name, N'' AS ShortName ");
                    sbSQLx.Append("UNION ALL ");
                    sbSQLx.Append("SELECT N'" + gCode + "' AS Code, N'' AS Name, N'' AS ShortName ");
                    sbSQLx.Append("ORDER BY Code, Name ");
                    new ObjDevEx.setGridLookUpEdit(glueCode, sbSQLx, "Code", "Code").getData(true);
                    glueCode.EditValue = gCode;
                }
            }


            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDCUST, Code, Name, ShortName, Contacts, Email, Address1, Address2, Address3, Country, PostCode, TelephoneNo, FaxNo, CustomerType, SalesSection, PaymentTerm, PaymentCurrency, CalendarNo, ");
            sbSQL.Append("       EvalutionPoint, OtherContact, OtherAddress1, OtherAddress2, OtherAddress3, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate ");
            sbSQL.Append("FROM   Customer ");
            sbSQL.Append("WHERE (Code = N'" + strCODE + "') ");
            string[] arrCust = new DBQuery(sbSQL).getMultipleValue();
            if (arrCust.Length > 0)
            {
                txeID.EditValue = arrCust[0];
                lblStatus.Text = "* Edit Customer";
                lblStatus.ForeColor = Color.Red;
                txeName.EditValue = arrCust[2];
                txeShortName.EditValue = arrCust[3];
                txeContacts.EditValue = arrCust[4];
                txeEmail.EditValue = arrCust[5];
                txeAddr1.EditValue = arrCust[6];
                txeAddr2.EditValue = arrCust[7];
                txeAddr3.EditValue = arrCust[8];
                txeCountry.EditValue = arrCust[9];
                txeTelNo.EditValue = arrCust[11];
                txeFaxNo.EditValue = arrCust[12];
                glueCustType.EditValue = arrCust[13];
                glueSection.EditValue = arrCust[14];
                glueTerm.EditValue = arrCust[15];
                glueCurrency.EditValue = arrCust[16];
                glueCalendar.EditValue = arrCust[17];
                txeEval.EditValue = arrCust[18]; ;
                txeOthContract.EditValue = arrCust[19];
                txeOthAddr1.EditValue = arrCust[20];
                txeOthAddr2.EditValue = arrCust[21];
                txeOthAddr3.EditValue = arrCust[22];
                txeCREATE.EditValue = arrCust[23];
                txeCDATE.EditValue = arrCust[24];
                txeUPDATE.EditValue = arrCust[25];
                txeUDATE.EditValue = arrCust[26];
            }
          
            
            //Check new customer or edit customer
            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCUST FROM Customer WHERE (OIDCUST = '" + txeID.EditValue.ToString() + "') ");
            string strCHKID = new DBQuery(sbSQL).getString();
            if (strCHKID == "")
            {
                lblStatus.Text = "* New Customer";
                lblStatus.ForeColor = Color.Green;
            }
            else
            {
                lblStatus.Text = "* Edit Customer";
                lblStatus.ForeColor = Color.Red;
            }

     
        }

        private void glueCode_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            
        }

        private void glueCode_ProcessNewValue(object sender, DevExpress.XtraEditors.Controls.ProcessNewValueEventArgs e)
        {

        }

        private void glueCode_Leave(object sender, EventArgs e)
        {
            ////glueCode.Text = glueCode.Text.ToUpper().Trim();
            //string gCode = glueCode.Text.ToUpper().Trim();
         
            //if (glueCode.Text != "")
            //{
            //    StringBuilder sbSQL = new StringBuilder();
            //    sbSQL.Append("SELECT OIDCUST FROM Customer WHERE (Code=N'" + gCode + "') ");
            //    string chkCode = new DBQuery(sbSQL).getString();
              
            //    if (chkCode == "")
            //    {
            //        sbSQL.Clear();
            //        sbSQL.Append("SELECT Code, Name, ShortName ");
            //        sbSQL.Append("FROM  Customer ");
            //        sbSQL.Append("UNION ALL ");
            //        sbSQL.Append("SELECT N'' AS Code, N'' AS Name, N'' AS ShortName ");
            //        sbSQL.Append("UNION ALL ");
            //        sbSQL.Append("SELECT N'" + gCode + "' AS Code, N'' AS Name, N'' AS ShortName ");
            //        sbSQL.Append("ORDER BY Code, Name ");
            //        new ObjDevEx.setGridLookUpEdit(glueCode, sbSQL, "Code", "Code").getData(true);
            //        glueCode.EditValue = gCode;
            //    }
            //}
        }

        private void gvCustomer_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            txeID.EditValue = gvCustomer.GetFocusedRowCellValue("No").ToString();
            lblStatus.Text = "* Edit Customer";
            lblStatus.ForeColor = Color.Red;
            glueCode.EditValue = gvCustomer.GetFocusedRowCellValue("Customer").ToString();
            txeName.EditValue = gvCustomer.GetFocusedRowCellValue("CustomerName").ToString();
            txeShortName.EditValue = gvCustomer.GetFocusedRowCellValue("ShortName").ToString();
            txeContacts.EditValue = gvCustomer.GetFocusedRowCellValue("ContactName").ToString();
            txeEmail.EditValue = gvCustomer.GetFocusedRowCellValue("Email").ToString();
            txeAddr1.EditValue = gvCustomer.GetFocusedRowCellValue("Address1").ToString();
            txeAddr2.EditValue = gvCustomer.GetFocusedRowCellValue("Address2").ToString();
            txeAddr3.EditValue = gvCustomer.GetFocusedRowCellValue("Address3").ToString();
            txeCountry.EditValue = gvCustomer.GetFocusedRowCellValue("Country").ToString();
            txeTelNo.EditValue = gvCustomer.GetFocusedRowCellValue("TelephoneNo").ToString();
            txeFaxNo.EditValue = gvCustomer.GetFocusedRowCellValue("FaxNo").ToString();
            glueCustType.EditValue = gvCustomer.GetFocusedRowCellValue("CustomerType").ToString();
            glueSection.EditValue = gvCustomer.GetFocusedRowCellValue("SalesSection").ToString();
            glueTerm.EditValue = gvCustomer.GetFocusedRowCellValue("PaymentTerm").ToString();
            glueCurrency.EditValue = gvCustomer.GetFocusedRowCellValue("PaymentCurrency").ToString();
            glueCalendar.EditValue = gvCustomer.GetFocusedRowCellValue("CalendarNo").ToString();
            txeEval.EditValue = gvCustomer.GetFocusedRowCellValue("CustomerEvalutionPoint").ToString();
            txeOthContract.EditValue = gvCustomer.GetFocusedRowCellValue("OtherContactName").ToString();
            txeOthAddr1.EditValue = gvCustomer.GetFocusedRowCellValue("OtherAddress1").ToString();
            txeOthAddr2.EditValue = gvCustomer.GetFocusedRowCellValue("OtherAddress2").ToString();
            txeOthAddr3.EditValue = gvCustomer.GetFocusedRowCellValue("OtherAddress3").ToString();

            txeCREATE.EditValue = gvCustomer.GetFocusedRowCellValue("CreatedBy").ToString();
            txeCDATE.EditValue = gvCustomer.GetFocusedRowCellValue("CreatedDate").ToString();
            txeUPDATE.EditValue = gvCustomer.GetFocusedRowCellValue("UpdatedBy").ToString();
            txeUDATE.EditValue = gvCustomer.GetFocusedRowCellValue("UpdatedDate").ToString();
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "CustomerList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvCustomer.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }
    }
}