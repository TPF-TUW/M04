﻿using System;
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
        private string selCode = "";
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
  
            //List<Product> products = new List<Product> {
            //    new Product(){ ProductName="Chang" },
            //    new Product(){ ProductName="Ipoh Coffee" },
            //    new Product(){ ProductName="Ravioli Angelo" },
            //    new Product(){ ProductName="Filo Mix" },
            //    new Product(){ ProductName="Tunnbröd" },
            //    new Product(){ ProductName="Konbu" },
            //    new Product(){ ProductName="Boston Crab Meat" }
            //};

            //glueCode.Properties.DataSource = products;
            glueCode.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            glueCode.Properties.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True;
            //glueCode.Properties.ValueMember = "ProductName";
            //glueCode.Properties.DisplayMember = glueCode.Properties.ValueMember;
            //glueCode.ProcessNewValue += glueCode_ProcessNewValue;

            bbiNew.PerformClick();
        }

        private void NewData()
        {
            txeID.EditValue = new DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCUST), '') = '' THEN 1 ELSE MAX(OIDCUST) + 1 END AS NewNo FROM Customer").getString();
            lblStatus.Text = "* Add Customer";
            lblStatus.ForeColor = Color.Green;
            glueCode.EditValue = "";
            txeName.Text = "";
            txeShortName.Text = "";
            txeContacts.Text = "";
            txeEmail.Text = "";
            txeAddr1.Text = "";
            txeAddr2.Text = "";
            txeAddr3.Text = "";
            txeCountry.Text = "";
            txePostCode.Text = "";
            txeTelNo.Text = "";
            txeFaxNo.Text = "";
            glueCustType.EditValue = "";
            glueSection.EditValue = "";
            glueTerm.EditValue = "";
            glueCurrency.EditValue = "";
            glueCalendar.EditValue = "";
            txeEval.Text = "";
            txeOthContract.Text = "";
            txeOthAddr1.Text = "";
            txeOthAddr2.Text = "";
            txeOthAddr3.Text = "";
            txeCREATE.Text = "0";
            txeCDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            txeUPDATE.Text = "0";
            txeUDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

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
            sbSQL.Append("SELECT CompanyName AS CustomerName, Year, OIDCALENDAR AS No  ");
            sbSQL.Append("FROM CalendarMaster A  ");
            sbSQL.Append("CROSS APPLY(SELECT ShortName AS CompanyName FROM Customer WHERE OIDCUST = A.OIDCompany) B  ");
            sbSQL.Append("WHERE CompanyType = 1  ");
            sbSQL.Append("ORDER BY Year DESC, CompanyName, OIDCALENDAR  ");
            new ObjDevEx.setGridLookUpEdit(glueCalendar, sbSQL, "CustomerName", "No").getData(true);

            //All Customer
            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCUST AS No, Code AS Customer, Name AS CustomerName, ShortName, Contacts AS ContactName, Email, Address1, Address2, Address3, Country, PostCode, TelephoneNo, ");
            sbSQL.Append("       FaxNo, CustomerType, SalesSection, PaymentTerm, PaymentCurrency, CalendarNo, EvalutionPoint AS CustomerEvalutionPoint, OtherContact AS OtherContactName, ");
            sbSQL.Append("       OtherAddress1, OtherAddress2, OtherAddress3, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate ");
            sbSQL.Append("FROM Customer ");
            sbSQL.Append("ORDER BY CustomerType, Code ");
            new ObjDevEx.setGridControl(gcCustomer, gvCustomer, sbSQL).getData(false, false, false, true);



        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            NewData();
            LoadData();
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (glueCode.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input customer code.");
                glueCode.Focus();
            }
            else if (txeName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input customer name.");
                txeName.Focus();
            }
            else if (txeShortName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input short name.");
                txeShortName.Focus();
            }
            else if (glueCustType.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select customer type.");
                glueCustType.Focus();
            }
            else if (glueCalendar.Text.Trim() == "")
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
                    if (txeCREATE.Text.Trim() != "")
                    {
                        strCREATE = txeCREATE.Text.Trim();
                    }

                    string strUPDATE = "0";
                    if (txeUPDATE.Text.Trim() != "")
                    {
                        strUPDATE = txeUPDATE.Text.Trim();
                    }

                    if (lblStatus.Text == "* Add Customer")
                    {
                        sbSQL.Append("  INSERT INTO Customer(Code, Name, ShortName, Contacts, Email, Address1, Address2, Address3, Country, PostCode, TelephoneNo, FaxNo, CustomerType, SalesSection, PaymentTerm, PaymentCurrency, CalendarNo, EvalutionPoint, OtherContact, OtherAddress1, OtherAddress2, OtherAddress3, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) ");
                        sbSQL.Append("  VALUES(N'" + glueCode.Text.Trim().Replace("'", "''") + "', N'" + txeName.Text.Trim().Replace("'", "''") + "', N'" + txeShortName.Text.Trim().Replace("'", "''") + "', N'" + txeContacts.Text.Trim().Replace("'", "''") + "', N'" + txeEmail.Text.Trim() + "', N'" + txeAddr1.Text.Trim() + "', N'" + txeAddr2.Text.Trim() + "', N'" + txeAddr3.Text.Trim() + "', N'" + txeCountry.Text.Trim() + "', N'" + txePostCode.Text.Trim() + "', N'" + txeTelNo.Text.Trim() + "', ");
                        sbSQL.Append("         N'" + txeFaxNo.Text.Trim() + "', '" + glueCustType.EditValue.ToString() + "', N'" + glueSection.Text.Trim() + "', N'" + glueTerm.Text.Trim() + "', N'" + glueCurrency.Text.Trim() + "', '" + glueCalendar.EditValue.ToString() + "', N'" + txeEval.Text.Trim() + "', N'" + txeOthContract.Text.Trim().Replace("'", "''") + "', N'" + txeOthAddr1.Text.Trim() + "', N'" + txeOthAddr2.Text.Trim() + "', N'" + txeOthAddr3.Text.Trim() + "', '" + strCREATE + "', GETDATE(), '" + strUPDATE + "', GETDATE()) ");
                    }
                    else if (lblStatus.Text == "* Edit Customer")
                    {
                        sbSQL.Append("  UPDATE Customer SET ");
                        sbSQL.Append("      Code = N'" + glueCode.Text.Trim().Replace("'", "''") + "', Name = N'" + txeName.Text.Trim().Replace("'", "''") + "', ShortName = N'" + txeShortName.Text.Trim().Replace("'", "''") + "', Contacts = N'" + txeContacts.Text.Trim().Replace("'", "''") + "', Email = N'" + txeEmail.Text.Trim() + "', Address1 = N'" + txeAddr1.Text.Trim() + "', Address2 = N'" + txeAddr2.Text.Trim() + "', Address3 = N'" + txeAddr3.Text.Trim() + "', ");
                        sbSQL.Append("      Country = N'" + txeCountry.Text.Trim() + "', PostCode = N'" + txePostCode.Text.Trim() + "', TelephoneNo = N'" + txeTelNo.Text.Trim() + "', FaxNo = N'" + txeFaxNo.Text.Trim() + "', CustomerType = '" + glueCustType.EditValue.ToString() + "', SalesSection = N'" + glueSection.Text.Trim() + "', PaymentTerm = N'" + glueTerm.Text.Trim() + "', ");
                        sbSQL.Append("      PaymentCurrency = N'" + glueCurrency.Text.Trim() + "', CalendarNo = '" + glueCalendar.EditValue.ToString() + "', EvalutionPoint = N'" + txeEval.Text.Trim() + "', OtherContact = N'" + txeOthContract.Text.Trim().Replace("'", "''") + "', OtherAddress1 = N'" + txeOthAddr1.Text.Trim() + "', OtherAddress2 = N'" + txeOthAddr2.Text.Trim() + "', OtherAddress3 = N'" + txeOthAddr3.Text.Trim() + "', ");
                        sbSQL.Append("      UpdatedBy = '" + strUPDATE + "', UpdatedDate = GETDATE() ");
                        sbSQL.Append("  WHERE(OIDCUST = '" + txeID.Text.Trim() + "') ");
                    }

                    //sbSQL.Append("IF NOT EXISTS(SELECT Code FROM Customer WHERE Code = N'" + glueCode.Text.Trim().Replace("'", "''") + "') ");
                    //sbSQL.Append(" BEGIN ");
                    //sbSQL.Append("  INSERT INTO Customer(Code, Name, ShortName, Contacts, Email, Address1, Address2, Address3, Country, PostCode, TelephoneNo, FaxNo, CustomerType, SalesSection, PaymentTerm, PaymentCurrency, CalendarNo, EvalutionPoint, OtherContact, OtherAddress1, OtherAddress2, OtherAddress3, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) ");
                    //sbSQL.Append("  VALUES(N'" + glueCode.Text.Trim().Replace("'", "''") + "', N'" + txeName.Text.Trim().Replace("'", "''") + "', N'" + txeShortName.Text.Trim().Replace("'", "''") + "', N'" + txeContacts.Text.Trim().Replace("'", "''") + "', N'" + txeEmail.Text.Trim() + "', N'" + txeAddr1.Text.Trim() + "', N'" + txeAddr2.Text.Trim() + "', N'" + txeAddr3.Text.Trim() + "', N'" + txeCountry.Text.Trim() + "', N'" + txePostCode.Text.Trim() + "', N'" + txeTelNo.Text.Trim() + "', ");
                    //sbSQL.Append("         N'" + txeFaxNo.Text.Trim() + "', '" + glueCustType.EditValue.ToString() + "', N'" + glueSection.Text.Trim() + "', N'" + glueTerm.Text.Trim() + "', N'" + glueCurrency.Text.Trim() + "', '" + glueCalendar.EditValue.ToString() + "', N'" + txeEval.Text.Trim() + "', N'" + txeOthContract.Text.Trim().Replace("'", "''") + "', N'" + txeOthAddr1.Text.Trim() + "', N'" + txeOthAddr2.Text.Trim() + "', N'" + txeOthAddr3.Text.Trim() + "', '" + strCREATE + "', GETDATE(), '" + strUPDATE + "', GETDATE()) ");
                    //sbSQL.Append(" END ");
                    //sbSQL.Append("ELSE ");
                    //sbSQL.Append(" BEGIN ");
                    //sbSQL.Append("  UPDATE Customer SET ");
                    //sbSQL.Append("      Code = N'" + glueCode.Text.Trim().Replace("'", "''") + "', Name = N'" + txeName.Text.Trim().Replace("'", "''") + "', ShortName = N'" + txeShortName.Text.Trim().Replace("'", "''") + "', Contacts = N'" + txeContacts.Text.Trim().Replace("'", "''") + "', Email = N'" + txeEmail.Text.Trim() + "', Address1 = N'" + txeAddr1.Text.Trim() + "', Address2 = N'" + txeAddr2.Text.Trim() + "', Address3 = N'" + txeAddr3.Text.Trim() + "', ");
                    //sbSQL.Append("      Country = N'" + txeCountry.Text.Trim() + "', PostCode = N'" + txePostCode.Text.Trim() + "', TelephoneNo = N'" + txeTelNo.Text.Trim() + "', FaxNo = N'" + txeFaxNo.Text.Trim() + "', CustomerType = '" + glueCustType.EditValue.ToString() + "', SalesSection = N'" + glueSection.Text.Trim() + "', PaymentTerm = N'" + glueTerm.Text.Trim() + "', ");
                    //sbSQL.Append("      PaymentCurrency = N'" + glueCurrency.Text.Trim() + "', CalendarNo = '" + glueCalendar.EditValue.ToString() + "', EvalutionPoint = N'" + txeEval.Text.Trim() + "', OtherContact = N'" + txeOthContract.Text.Trim().Replace("'", "''") + "', OtherAddress1 = N'" + txeOthAddr1.Text.Trim() + "', OtherAddress2 = N'" + txeOthAddr2.Text.Trim() + "', OtherAddress3 = N'" + txeOthAddr3.Text.Trim() + "', ");
                    //sbSQL.Append("      UpdatedBy = '" + strUPDATE +"', UpdatedDate = GETDATE() ");
                    //sbSQL.Append("  WHERE(OIDCUST = '" + txeID.Text.Trim() + "') ");
                    //sbSQL.Append(" END ");
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
            //txeName.Focus();
            //LoadCode(glueCode.Text);
        }


        private void glueCodeX_EditValueChanged(object sender, EventArgs e)
        {
            //Display lookup editor's current value.
            //LookUpEditBase lookupEditor = sender as LookUpEditBase;
            //if (lookupEditor == null) return;
            
            //if (lookupEditor.EditValue == null)
            //    layoutControlItem17.Text = "Current EditValue: null";
            //else
            //    layoutControlItem17.Text = "Current EditValue: " + lookupEditor.EditValue.ToString();
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
            if (glueCode.Text.Trim() != "" && glueCode.Text.ToUpper().Trim() != selCode)
            {
                glueCode.Text = glueCode.Text.ToUpper().Trim();
                selCode = glueCode.Text;
                LoadCode(glueCode.Text);  
            }
            
        }

        private void LoadCode(string strCODE)
        {
            strCODE = strCODE.ToUpper().Trim();
            txeID.Text = "";
            lblStatus.Text = "* Add Customer";
            lblStatus.ForeColor = Color.Green;
            txeName.Text = "";
            txeShortName.Text = "";
            txeContacts.Text = "";
            txeEmail.Text = "";
            txeAddr1.Text = "";
            txeAddr2.Text = "";
            txeAddr3.Text = "";
            txeCountry.Text = "";
            txePostCode.Text = "";
            txeTelNo.Text = "";
            txeFaxNo.Text = "";
            glueCustType.EditValue = "";
            glueSection.EditValue = "";
            glueTerm.EditValue = "";
            glueCurrency.EditValue = "";
            glueCalendar.EditValue = "";
            txeEval.Text = "";
            txeOthContract.Text = "";
            txeOthAddr1.Text = "";
            txeOthAddr2.Text = "";
            txeOthAddr3.Text = "";
            txeCREATE.Text = "0";
            txeCDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            txeUPDATE.Text = "0";
            txeUDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDCUST, Code, Name, ShortName, Contacts, Email, Address1, Address2, Address3, Country, PostCode, TelephoneNo, FaxNo, CustomerType, SalesSection, PaymentTerm, PaymentCurrency, CalendarNo, ");
            sbSQL.Append("       EvalutionPoint, OtherContact, OtherAddress1, OtherAddress2, OtherAddress3, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate ");
            sbSQL.Append("FROM   Customer ");
            sbSQL.Append("WHERE (Code = N'" + strCODE.Replace("'", "''") + "') ");
            string[] arrCust = new DBQuery(sbSQL).getMultipleValue();
            if (arrCust.Length > 0)
            {
                txeID.Text = arrCust[0];
                lblStatus.Text = "* Edit Customer";
                lblStatus.ForeColor = Color.Red;
                txeName.Text = arrCust[2];
                txeShortName.Text = arrCust[3];
                txeContacts.Text = arrCust[4];
                txeEmail.Text = arrCust[5];
                txeAddr1.Text = arrCust[6];
                txeAddr2.Text = arrCust[7];
                txeAddr3.Text = arrCust[8];
                txeCountry.Text = arrCust[9];
                txePostCode.Text = arrCust[10];
                txeTelNo.Text = arrCust[11];
                txeFaxNo.Text = arrCust[12];
                glueCustType.EditValue = arrCust[13];
                glueSection.EditValue = arrCust[14];
                glueTerm.EditValue = arrCust[15];
                glueCurrency.EditValue = arrCust[16];
                glueCalendar.EditValue = arrCust[17];
                txeEval.Text = arrCust[18]; ;
                txeOthContract.Text = arrCust[19];
                txeOthAddr1.Text = arrCust[20];
                txeOthAddr2.Text = arrCust[21];
                txeOthAddr3.Text = arrCust[22];
                txeCREATE.Text = arrCust[23];
                txeCDATE.Text = arrCust[24];
                txeUPDATE.Text = arrCust[25];
                txeUDATE.Text = arrCust[26];
            }


            //Check new customer or edit customer
            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCUST FROM Customer WHERE (OIDCUST = '" + txeID.EditValue.ToString() + "') ");
            string strCHKID = new DBQuery(sbSQL).getString();
            if (strCHKID == "")
            {
                lblStatus.Text = "* Add Customer";
                lblStatus.ForeColor = Color.Green;
            }
            else
            {
                lblStatus.Text = "* Edit Customer";
                lblStatus.ForeColor = Color.Red;
            }
            txeName.Focus();
        }

        private void gvCustomer_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            txeID.Text = gvCustomer.GetFocusedRowCellValue("No").ToString();
            lblStatus.Text = "* Edit Customer";
            lblStatus.ForeColor = Color.Red;
            glueCode.EditValue = gvCustomer.GetFocusedRowCellValue("Customer").ToString();
            txeName.Text = gvCustomer.GetFocusedRowCellValue("CustomerName").ToString();
            txeShortName.Text = gvCustomer.GetFocusedRowCellValue("ShortName").ToString();
            txeContacts.Text = gvCustomer.GetFocusedRowCellValue("ContactName").ToString();
            txeEmail.Text = gvCustomer.GetFocusedRowCellValue("Email").ToString();
            txeAddr1.Text = gvCustomer.GetFocusedRowCellValue("Address1").ToString();
            txeAddr2.Text = gvCustomer.GetFocusedRowCellValue("Address2").ToString();
            txeAddr3.Text = gvCustomer.GetFocusedRowCellValue("Address3").ToString();
            txeCountry.Text = gvCustomer.GetFocusedRowCellValue("Country").ToString();
            txePostCode.Text = gvCustomer.GetFocusedRowCellValue("PostCode").ToString();
            txeTelNo.Text = gvCustomer.GetFocusedRowCellValue("TelephoneNo").ToString();
            txeFaxNo.Text = gvCustomer.GetFocusedRowCellValue("FaxNo").ToString();
            glueCustType.EditValue = gvCustomer.GetFocusedRowCellValue("CustomerType").ToString();
            glueSection.EditValue = gvCustomer.GetFocusedRowCellValue("SalesSection").ToString();
            glueTerm.EditValue = gvCustomer.GetFocusedRowCellValue("PaymentTerm").ToString();
            glueCurrency.EditValue = gvCustomer.GetFocusedRowCellValue("PaymentCurrency").ToString();
            glueCalendar.EditValue = gvCustomer.GetFocusedRowCellValue("CalendarNo").ToString();
            txeEval.Text = gvCustomer.GetFocusedRowCellValue("CustomerEvalutionPoint").ToString();
            txeOthContract.Text = gvCustomer.GetFocusedRowCellValue("OtherContactName").ToString();
            txeOthAddr1.Text = gvCustomer.GetFocusedRowCellValue("OtherAddress1").ToString();
            txeOthAddr2.Text = gvCustomer.GetFocusedRowCellValue("OtherAddress2").ToString();
            txeOthAddr3.Text = gvCustomer.GetFocusedRowCellValue("OtherAddress3").ToString();

            txeCREATE.Text = gvCustomer.GetFocusedRowCellValue("CreatedBy").ToString();
            txeCDATE.Text = gvCustomer.GetFocusedRowCellValue("CreatedDate").ToString();
            txeUPDATE.Text = gvCustomer.GetFocusedRowCellValue("UpdatedBy").ToString();
            txeUDATE.Text = gvCustomer.GetFocusedRowCellValue("UpdatedDate").ToString();
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "CustomerList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvCustomer.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void txeName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeShortName.Focus();
            }
        }

        private void txeShortName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeContacts.Focus();
            }
        }

        private void txeContacts_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeEmail.Focus();
            }
        }

        private void txeEmail_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeAddr1.Focus();
            }
        }

        private void txeAddr1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeAddr2.Focus();
            }
        }

        private void txeAddr2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeAddr3.Focus();
            }
        }

        private void txeAddr3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeCountry.Focus();
            }
        }

        private void txeCountry_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txePostCode.Focus();
            }
        }

        private void txePostCode_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode==Keys.Enter)
            {
                txeTelNo.Focus();
            }
        }

        private void txeTelNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeFaxNo.Focus();
            }
        }

        private void txeFaxNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                glueCustType.Focus();
            }
        }

        private void glueCustType_EditValueChanged(object sender, EventArgs e)
        {
            glueSection.Focus();
        }

        private void glueSection_EditValueChanged(object sender, EventArgs e)
        {
            glueTerm.Focus();
        }

        private void glueTerm_EditValueChanged(object sender, EventArgs e)
        {
            glueCurrency.Focus();
        }

        private void glueCurrency_EditValueChanged(object sender, EventArgs e)
        {
            glueCalendar.Focus();
        }

        private void glueCalendar_EditValueChanged(object sender, EventArgs e)
        {
            txeEval.Focus();
        }

        private void txeEval_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeOthContract.Focus();
            }
        }

        private void txeOthContract_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeOthAddr1.Focus();
            }
        }

        private void txeOthAddr1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeOthAddr2.Focus();
            }
        }

        private void txeOthAddr2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeOthAddr3.Focus();
            }
        }

        private void gvCustomer_RowStyle(object sender, RowStyleEventArgs e)
        {
            if (sender is GridView)
            {
                GridView gView = (GridView)sender;
                if (!gView.IsValidRowHandle(e.RowHandle)) return;
                int parent = gView.GetParentRowHandle(e.RowHandle);
                if (gView.IsGroupRow(parent))
                {
                    for (int i = 0; i < gView.GetChildRowCount(parent); i++)
                    {
                        if (gView.GetChildRowHandle(parent, i) == e.RowHandle)
                        {
                            e.Appearance.BackColor = i % 2 == 0 ? Color.AliceBlue : Color.White;
                        }
                    } 
                }
                else
                {
                    e.Appearance.BackColor = e.RowHandle % 2 == 0 ? Color.AliceBlue : Color.White;
                }
            }
        }

        private void glueCode_ProcessNewValue(object sender, DevExpress.XtraEditors.Controls.ProcessNewValueEventArgs e)
        {
            GridLookUpEdit gridLookup = sender as GridLookUpEdit;
            if (e.DisplayValue == null) return;
            string newValue = e.DisplayValue.ToString();
            if (newValue == String.Empty) return;
        }

        private void glueCode_CloseUp(object sender, DevExpress.XtraEditors.Controls.CloseUpEventArgs e)
        {
            
        }

        private void glueCode_Closed(object sender, DevExpress.XtraEditors.Controls.ClosedEventArgs e)
        {
            //glueCode.Text = glueCode.Text.ToUpper().Trim();
            //LoadCode(glueCode.Text);
           // MessageBox.Show(glueCode.Text);
            glueCode.Focus();
            txeName.Focus();
        }
    }
}