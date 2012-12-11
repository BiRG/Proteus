using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.IO;

public partial class _Default : Page
{
    public int loadcount = 0;
    public static String connectionString = "Data Source=geodata.cofc.edu;Initial Catalog=Avkat_mysql;Integrated Security=SSPI";
    public String FILE_NAME = @"C:\QueryOutputs\QueryResult.csv";
    public static string photoQueryString = "http://earth.cofc.edu/avkat_photos/viewphotos.html?";

    public static string sendPicInfo(string id, string name)
    {
        List<string> photos = new List<string>();
        string photoString = "";
        string listString = "[";
        string sql = string.Format("select photoID from Photolog where FUSUID = '{0}'", id);
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            SqlCommand com = new SqlCommand(sql, connection);
            connection.Open();
            SqlDataReader reader = com.ExecuteReader();
            while (reader.Read())
            {
                string photoid = reader["photoID"].ToString();
                photos.Add(photoid);
            }
            connection.Close();
            photoString = "FeatureName=" + name + "&" + "FeatureID=" + id + "&PhotoID=";

            foreach (string i in photos)
            {
                listString += i + ",";
            }
            listString += "]";
            photoString += listString;
            photoQueryString += photoString;
        }
        return "";
    }

    

    protected void Page_Init(object sender, EventArgs e)
    {
        DropDownList1.Items.Add("Feature");
        DropDownList1.Items.Add("Survey Unit");
        DropDownList1.Items.Add("Ceramic");
        DropDownList1.Items.Add("Image");
            

        String sql = "select Period_code from chronology";
        String sql2 = "select distinct FeatureType from features";
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            SqlCommand com = new SqlCommand(sql, connection);
            SqlCommand com2 = new SqlCommand(sql2, connection);
            connection.Open();
            SqlDataReader reader = com.ExecuteReader();
                
            while (reader.Read())
            {
                DropDownList3.Items.Add(reader["Period_code"].ToString());
            }
            connection.Close();
            connection.Open();
            SqlDataReader reader2 = com2.ExecuteReader();
            while (reader2.Read())
            {
                DropDownList2.Items.Add(reader2["FeatureType"].ToString());
            }
            connection.Close();
        }
    }
    
    protected void Page_Load(object sender, EventArgs e)
    {
        
    }
        

       


    protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!DropDownList1.Text.Equals("Feature"))
        {
            DropDownList2.Enabled = false;
        }
        else
        {
            DropDownList2.Enabled = true;
        }
    }
    protected void DropDownList2_SelectedIndexChanged(object sender, EventArgs e)
    {

    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        if (DropDownList1.Text.Equals("Feature"))
        {
            bool useEra = false;
            if(TextBox1.Text.Equals("") || TextBox2.Text.Equals(""))
            {
                useEra = true;
            }
            String sql = string.Format("select FeatureName, FeatureType, FeatureID, Begin_Date, End_Date, Ceramics from features where FeatureType = '{0}' and {1} != 0 ", DropDownList2.Text, DropDownList3.Text);
            if (!useEra)
            {
                sql = string.Format("select FeatureName, FeatureType, FeatureID, Begin_Date, End_Date, Ceramics from features where FeatureType = '{0}' and Begin_date >= {1} and End_date <= {2}", DropDownList2.Text, TextBox1.Text, TextBox2.Text);
            }

            TableRow initRow = new TableRow();
            Table1.Rows.Add(initRow);
            TableCell idCell1 = new TableCell();
            TableCell fnCell = new TableCell();
            TableCell ftCell = new TableCell();
            TableCell begCell = new TableCell();
            TableCell endCell = new TableCell();
            TableCell photoCell1 = new TableCell();
            TableCell but1Cell = new TableCell();
            TableCell cerCell = new TableCell();
            
            photoCell1.Text = "View Photo";
            but1Cell.Text = "View Detailed info";
            idCell1.Text ="Feature id";
            fnCell.Text = "Feature Name";
            ftCell.Text = "feature Type";
            begCell.Text = "Date begin";
            endCell.Text = "Date End";
          

            initRow.Cells.Add(but1Cell);
            initRow.Cells.Add(photoCell1);
            initRow.Cells.Add(idCell1);
            initRow.Cells.Add(fnCell);
            initRow.Cells.Add(ftCell);
            initRow.Cells.Add(begCell);
            initRow.Cells.Add(endCell);
          
            
                
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand com = new SqlCommand(sql, connection);
                connection.Open();
                SqlDataReader reader = com.ExecuteReader();
                while (reader.Read())
                {
                    string Fname = reader["FeatureName"].ToString();
                    string ftype = reader["FeatureType"].ToString();
                    string fid = reader["FeatureID"].ToString();
                    string dateBegin = reader["Begin_date"].ToString();
                    string dateEnd = reader["End_Date"].ToString();
                    TableRow tRow = new TableRow();
                    Table1.Rows.Add(tRow);
                    TableCell fCell = new TableCell();
                    TableCell tCell = new TableCell();
                    TableCell butCell = new TableCell();
                    TableCell idCell = new TableCell();
                    TableCell bdCell = new TableCell();
                    TableCell edCell = new TableCell();
                    TableCell photoCell = new TableCell();
                    fCell.Text = Fname;
                    tCell.Text = ftype;
                    bdCell.Text = dateBegin;
                    edCell.Text = dateEnd;
                    idCell.Text = fid;
                    butCell.Controls.Add(CreateButton(fid, Fname));
                    photoCell.Controls.Add(CreatePhotoButton(fid, Fname));
                    tRow.Cells.Add(butCell);
                    tRow.Cells.Add(photoCell);
                    tRow.Cells.Add(idCell);
                    tRow.Cells.Add(fCell);
                    tRow.Cells.Add(tCell);
                    tRow.Cells.Add(bdCell);
                    tRow.Cells.Add(edCell);


                }
                connection.Close();
            }
        }

            else if (DropDownList1.Text.Equals("Survey Unit"))// data for survey units is not yet functional
            {
                
                String sql = string.Format("select SUID, Date, Ceramics from surveyunits where Ceramics = {0}", DropDownList4.Text);

                TableRow initRow = new TableRow();
                Table1.Rows.Add(initRow);
                TableCell s1Cell = new TableCell();
                TableCell d1Cell = new TableCell();
                TableCell c1Cell = new TableCell();
                TableCell but1Cell = new TableCell();
                TableCell photoCell1 = new TableCell();
                photoCell1.Text = "View Photo";
                but1Cell.Text = "View Detailed info";
                s1Cell.Text = "Survey Unit";
                d1Cell.Text = "Date";
                c1Cell.Text = "Ceramic(Y/N)";

                initRow.Cells.Add(but1Cell);
                initRow.Cells.Add(photoCell1);
                initRow.Cells.Add(s1Cell);
                initRow.Cells.Add(c1Cell);
                initRow.Cells.Add(d1Cell);
                
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand com = new SqlCommand(sql, connection);
                    connection.Open();
                    SqlDataReader reader = com.ExecuteReader();
                    while (reader.Read())
                    {
                        string suid = reader["SUID"].ToString();
                        string date = reader["Date"].ToString();
                        string cer = reader["Ceramics"].ToString();
                       
                        TableRow tRow = new TableRow();
                        Table1.Rows.Add(tRow);
                        TableCell sCell = new TableCell();
                        TableCell dCell = new TableCell();
                        TableCell cCell = new TableCell();
                        TableCell butCell = new TableCell();
                        TableCell photoCell = new TableCell();
                        sCell.Text = suid;
                        dCell.Text = date;
                        cCell.Text = cer;
                       
                        butCell.Controls.Add(CreateButton(suid, "Survey Unit"));
                        photoCell.Controls.Add(CreatePhotoButton(suid,"Survey Unit"));
                        tRow.Cells.Add(butCell);
                        tRow.Cells.Add(photoCell);
                        tRow.Cells.Add(sCell);
                        tRow.Cells.Add(cCell);
                        tRow.Cells.Add(dCell);
                        
                    }
                }   
            }
        else if(DropDownList1.Text.Equals("Ceramic"))
        {
            bool useEra = false;
            if (TextBox1.Text.Equals("") || TextBox2.Text.Equals(""))
            {
                useEra = true;
            }
            String sql = string.Format("select Unit, Inventory_no, Vessel_Type_Count, Begin_date, End_date, Rim_Count, Body_Count, Handle_Count, Base_Count, Lid_Count, Spout_Count, Complete_Count, Tile_Count, Shape_Count, Pithos_Count, Plain_Storage_Count, Cooking_Count, Plain_Open_Count, Plain_Closed_Count, Fine_Count, Glazed_Cooking_Count, Glazed_Open_Count, Glazed_Closed_Count, Unknown_Type_Count, Other_Ceramic_Type, Other_Type_Description from v_ceramics where {0} != 0 ", DropDownList2.Text);
            if (!useEra)
            {
                sql = string.Format("select Unit, Inventory_no, Vessel_Type_Count, Begin_date, End_date, Rim_Count, Body_Count, Handle_Count, Base_Count, Lid_Count, Spout_Count, Complete_Count, Tile_Count, Shape_Count, Pithos_Count, Plain_Storage_Count, Cooking_Count, Plain_Open_Count, Plain_Closed_Count, Fine_Count, Glazed_Cooking_Count, Glazed_Open_Count, Glazed_Closed_Count, Unknown_Type_Count, Other_Ceramic_Type, Other_Type_Description from v_ceramics where Begin_Date >= {0} and End_Date <= {1}  ", TextBox1.Text, TextBox2.Text);
            }

            TableRow initRow = new TableRow();
            Table1.Rows.Add(initRow);
            TableCell photoCell1 = new TableCell();
            TableCell but1Cell = new TableCell();
            TableCell uCelli = new TableCell();
            TableCell InCelli = new TableCell();
            TableCell vtCelli = new TableCell();
            TableCell rcCelli = new TableCell();
            TableCell bodyCell1 = new TableCell();
            TableCell hcCelli = new TableCell();
            TableCell bcCelli = new TableCell();
            TableCell lcCelli = new TableCell();
            TableCell scCelli = new TableCell();
            TableCell ccCelli = new TableCell();
            TableCell tileCelli = new TableCell();
            TableCell shapecCelli = new TableCell();
            TableCell pcCelli = new TableCell();
            TableCell pscCelli = new TableCell();
            TableCell cookCelli = new TableCell();
            TableCell plainoCelli = new TableCell();
            TableCell plaincCelli = new TableCell();
            TableCell fineCelli = new TableCell();
            TableCell glazedcCelli = new TableCell();
            TableCell glazedOCelli = new TableCell();
            TableCell glazedCCelli = new TableCell();
            TableCell unkCelli = new TableCell();
            TableCell otherCelli = new TableCell();
            TableCell otherdCelli = new TableCell();



            photoCell1.Text = "View Photo";
            but1Cell.Text = "View Detailed info";
            uCelli.Text = "Unit";
            InCelli.Text = "Inventory#";
            vtCelli.Text = "Vessel Type Count";
            rcCelli.Text = "Rim Count";
            bodyCell1.Text = "Body Count";
            hcCelli.Text = "Handle Count";
            bcCelli.Text = "Base Count";
            lcCelli.Text = "Lid Count";
            scCelli.Text = "Spout Count";
            ccCelli.Text = "Complete Count";
            tileCelli.Text = "Tile Count";
            shapecCelli.Text = "Shape Count";
            pcCelli.Text = "Pithos Count";
            pscCelli.Text = "Plain_Storage_Count";
            cookCelli.Text = "Cooking Count";
            plainoCelli.Text = "Plain Open Count";
            plaincCelli.Text = "Plain Closed Count";
            fineCelli.Text = "Fine Count";
            glazedcCelli.Text = "Glazed Cooking Count";
            glazedOCelli.Text = "Glazed Open Count";
            glazedCCelli.Text = "Glazed Closed Count";
            unkCelli.Text = "Unknown Count";
            otherCelli.Text = "Other Type Count";
            otherdCelli.Text = "Other Type Description";
            


            initRow.Cells.Add(but1Cell);
            initRow.Cells.Add(photoCell1);
            initRow.Cells.Add(uCelli);
            initRow.Cells.Add(InCelli);
            initRow.Cells.Add(vtCelli);
            initRow.Cells.Add(rcCelli);
            initRow.Cells.Add(bodyCell1);
            initRow.Cells.Add(hcCelli);
            initRow.Cells.Add(bcCelli);
            initRow.Cells.Add(lcCelli);
            initRow.Cells.Add(scCelli);
            initRow.Cells.Add(ccCelli);
            initRow.Cells.Add(tileCelli);
            initRow.Cells.Add(shapecCelli);
            initRow.Cells.Add(pcCelli);
            initRow.Cells.Add(pscCelli);
            initRow.Cells.Add(cookCelli);
            initRow.Cells.Add(plainoCelli);
            initRow.Cells.Add(plaincCelli);
            initRow.Cells.Add(fineCelli);
            initRow.Cells.Add(glazedcCelli);
            initRow.Cells.Add(glazedOCelli);
            initRow.Cells.Add(glazedCCelli);
            initRow.Cells.Add(unkCelli);
            initRow.Cells.Add(otherCelli);
            initRow.Cells.Add(otherdCelli);


            

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand com = new SqlCommand(sql, connection);
                connection.Open();
                SqlDataReader reader = com.ExecuteReader();
                while (reader.Read())
                {
                    string unit = reader["Unit"].ToString();
                    string inventory = reader["Inventory_no"].ToString();
                    string bdate = reader["Begin_date"].ToString();
                    string edate = reader["End_date"].ToString();
                    string rimc = reader["Rim_Count"].ToString();
                    string bodyc = reader["Body_Count"].ToString();
                    string handlec = reader["Handle_Count"].ToString();
                    string basec = reader["Base_Count"].ToString();
                    string lidc = reader["Lid_Count"].ToString();
                    string spoutc = reader["Spout_Count"].ToString();
                    string compc = reader["Complete_Count"].ToString();
                    string tilec = reader["Tile_Count"].ToString();
                    string shapec = reader["Shape_Count"].ToString();
                    string pithosc = reader["Pithos_Count"].ToString();
                    string plainc = reader["Plain_Storage_Count"].ToString();
                    string cookc = reader["Cooking_Count"].ToString();
                    string plainoc = reader["Plain_Open_Count"].ToString();
                    string plaincc = reader["Plain_Closed_Count"].ToString();
                    string finec = reader["Fine_Count"].ToString();
                    string glazedcc = reader["Glazed_Cooking_Count"].ToString();
                    string glazedoc = reader["Glazed_Open_Count"].ToString();
                    string glazedclc = reader["Glazed_Closed_Count"].ToString();
                    string unkc = reader["Unknown_Type_Count"].ToString();
                    string otherc = reader["Other_Ceramic_Type"].ToString();
                    string otherd = reader["Other_Type_Description"].ToString();
                    string vesselc = reader["Vessel_Type_Count"].ToString();

                    TableRow tRow = new TableRow();
                    Table1.Rows.Add(tRow);
                    TableCell photoCell = new TableCell();
                    TableCell butCell = new TableCell();
                    TableCell uCell = new TableCell();
                    TableCell InCell = new TableCell();
                    TableCell vtCell = new TableCell();
                    TableCell rcCell = new TableCell();
                    TableCell bodyCell = new TableCell();
                    TableCell hcCell = new TableCell();
                    TableCell bcCell = new TableCell();
                    TableCell lcCell = new TableCell();
                    TableCell scCell = new TableCell();
                    TableCell ccCell = new TableCell();
                    TableCell tileCell = new TableCell();
                    TableCell shapecCell = new TableCell();
                    TableCell pcCell = new TableCell();
                    TableCell pscCell = new TableCell();
                    TableCell cookCell = new TableCell();
                    TableCell plainoCell = new TableCell();
                    TableCell plaincCell = new TableCell();
                    TableCell fineCell = new TableCell();
                    TableCell glazedcCell = new TableCell();
                    TableCell glazedOCell = new TableCell();
                    TableCell glazedCCell = new TableCell();
                    TableCell unkCell = new TableCell();
                    TableCell otherCell = new TableCell();
                    TableCell otherdCell = new TableCell();

                    butCell.Controls.Add(CreateButton(unit, "Ceramic"));
                    photoCell.Controls.Add(CreatePhotoButton(unit, "Ceramic"));
                    uCell.Text = unit;
                    InCell.Text = inventory;
                    vtCell.Text = vesselc;
                    rcCell.Text = rimc;
                    bodyCell.Text = bodyc;
                    hcCell.Text = handlec;
                    bcCell.Text = basec;
                    lcCell.Text = lidc;
                    scCell.Text = spoutc;
                    ccCell.Text = compc;
                    tileCell.Text = tilec;
                    shapecCell.Text = shapec;
                    pcCell.Text = pithosc;
                    pscCell.Text = plainc;
                    cookCell.Text = cookc;
                    plainoCell.Text = plainoc;
                    plaincCell.Text = plaincc;
                    fineCell.Text = finec;
                    glazedcCell.Text = glazedcc;
                    glazedOCell.Text = glazedoc;
                    glazedCCell.Text = glazedclc;
                    unkCell.Text = unkc;
                    otherCell.Text = otherc;
                    otherdCell.Text = otherd;

                    tRow.Cells.Add(photoCell);
                    tRow.Cells.Add(butCell);
                    tRow.Cells.Add(uCell);
                    tRow.Cells.Add(InCell);
                    tRow.Cells.Add(vtCell);
                    tRow.Cells.Add(rcCell);
                    tRow.Cells.Add(bodyCell);
                    tRow.Cells.Add(hcCell);
                    tRow.Cells.Add(bcCell);
                    tRow.Cells.Add(lcCell);
                    tRow.Cells.Add(scCell);
                    tRow.Cells.Add(ccCell);
                    tRow.Cells.Add(tileCell);
                    tRow.Cells.Add(shapecCell);
                    tRow.Cells.Add(pcCell);
                    tRow.Cells.Add(pscCell);
                    tRow.Cells.Add(cookCell);
                    tRow.Cells.Add(plainoCell);
                    tRow.Cells.Add(plaincCell);
                    tRow.Cells.Add(fineCell);
                    tRow.Cells.Add(glazedcCell);
                    tRow.Cells.Add(glazedOCell);
                    tRow.Cells.Add(glazedCCell);
                    tRow.Cells.Add(unkCell);
                    tRow.Cells.Add(otherCell);
                    tRow.Cells.Add(otherdCell);



                }
            }
        }
    }


    private Button CreateButton(string id, string name)
    {
        Button b = new Button();
        List<string> photos = new List<string>();
        b.Text = "info";
        b.ID = id;
       //b.OnClick() = new ; // this is where we would add a method that links to the legacy page for further information
        return b;                                              //on a survey unit or feature
    }

    private Button CreatePhotoButton(string id, string name)
    {
        
        Button b = new Button();
        b.Text = "photo";
        b.ID = id;
        
       
        sendPicInfo(id, name);
        string s = photoQueryString;
        b.OnClientClick = "Navigate()";
        //b.OnClientClick = "Navigate()" + sendPicInfo(id,name);
       // b.Click +=  new EventHandler(this.sendPicInfo);// this is where we would add a method that links to the legacy page for further information
        return b;                                              //on a survey unit or feature
    }

    private void goToLegacy(String featureid) // open a new tab to our legacy site to view further information
    {

    }
    

    protected void Button2_Click(object sender, EventArgs e)
    {
       
        if (DropDownList1.Text.Equals("Feature"))
        {
            bool useEra = false;
            if (TextBox1.Text.Equals("") || TextBox2.Text.Equals(""))
            {
                useEra = true;
            }
            String sql = string.Format("select FeatureName, FeatureType, FeatureID, Begin_Date, End_Date, Ceramics from features where FeatureType = '{0}' and {1} != 0 ", DropDownList2.Text, DropDownList3.Text);
            if (!useEra)
            {
                sql = string.Format("select FeatureName, FeatureType, FeatureID, Begin_Date, End_Date, Ceramics from features where FeatureType = '{0}' and Begin_date >= {1} and End_date <= {2}", DropDownList2.Text, TextBox1.Text, TextBox2.Text);
            }

            using (FileStream filestream = new FileStream(FILE_NAME, FileMode.CreateNew))
            {
                
                StreamWriter writer = new StreamWriter(filestream);
                writer.WriteLine("FeatureName,FeatureType,FeatureID,BeginDate,EndDate");
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand com = new SqlCommand(sql, connection);
                    connection.Open();
                    SqlDataReader reader = com.ExecuteReader();
                    while (reader.Read())
                    {
                        string Fname = reader["FeatureName"].ToString();
                        string ftype = reader["FeatureType"].ToString();
                        string fid = reader["FeatureID"].ToString();
                        string dateBegin = reader["Begin_date"].ToString();
                        string dateEnd = reader["End_Date"].ToString();
                        writer.WriteLine(Fname + "," + ftype + "," + fid + "," + dateBegin + "," + dateEnd);

                        
                    }
                    connection.Close();
                    writer.Close();
                }
            }
            Response.ContentType = "application/csv";
            Response.AppendHeader("Content-Disposition", "attachment; filename=QueryResult.csv");
            string FilePath = FILE_NAME;
            Response.TransmitFile(FilePath);
            Response.End();
        }

        else if (DropDownList1.Text.Equals("Survey Unit"))// data for survey units is not yet functional
        {

            String sql = string.Format("select SUID, Date, Ceramics from surveyunits where Ceramics = {0}", DropDownList4.Text);
            using (FileStream filestream = new FileStream(FILE_NAME, FileMode.CreateNew))
            {
                StreamWriter writer = new StreamWriter(filestream);
                writer.WriteLine("SUID,Date,Ceramics(Y/N)");
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand com = new SqlCommand(sql, connection);
                    connection.Open();
                    SqlDataReader reader = com.ExecuteReader();
                    while (reader.Read())
                    {
                        string suid = reader["SUID"].ToString();
                        string date = reader["Date"].ToString();
                        string cer = reader["Ceramics"].ToString();
                        writer.WriteLine(suid + "," + date + "," + cer);


                    }
                    connection.Close();
                    writer.Close();
                }
            }
            Response.ContentType = "application/csv";
            Response.AppendHeader("Content-Disposition", "attachment; filename=QueryResult.csv");
            string FilePath = FILE_NAME;
            Response.TransmitFile(FilePath);
            Response.End();
        }

        else if (DropDownList1.Text.Equals("Ceramic"))
        {
            bool useEra = false;
            if (TextBox1.Text.Equals("") || TextBox2.Text.Equals(""))
            {
                useEra = true;
            }
            String sql = string.Format("select Unit, Inventory_no, Vessel_Type_Count, Begin_date, End_date, Rim_Count, Body_Count, Handle_Count, Base_Count, Lid_Count, Spout_Count, Complete_Count, Tile_Count, Shape_Count, Pithos_Count, Plain_Storage_Count, Cooking_Count, Plain_Open_Count, Plain_Closed_Count, Fine_Count, Glazed_Cooking_Count, Glazed_Open_Count, Glazed_Closed_Count, Unknown_Type_Count, Other_Ceramic_Type, Other_Type_Description from v_ceramics where {0} != 0 ", DropDownList2.Text);
            if (!useEra)
            {
                sql = string.Format("select Unit, Inventory_no, Vessel_Type_Count, Begin_date, End_date, Rim_Count, Body_Count, Handle_Count, Base_Count, Lid_Count, Spout_Count, Complete_Count, Tile_Count, Shape_Count, Pithos_Count, Plain_Storage_Count, Cooking_Count, Plain_Open_Count, Plain_Closed_Count, Fine_Count, Glazed_Cooking_Count, Glazed_Open_Count, Glazed_Closed_Count, Unknown_Type_Count, Other_Ceramic_Type, Other_Type_Description from v_ceramics where Begin_Date >= {0} and End_Date <= {1}  ", TextBox1.Text, TextBox2.Text);
            }

            using (FileStream filestream = new FileStream(FILE_NAME, FileMode.CreateNew))
            {
                StreamWriter writer = new StreamWriter(filestream);
                writer.WriteLine("Unit,Inventory#,Vessel Type Count,Begin Date,End Date,Rim Count,Body Count,Handle Count,Base Count,Lid Count,Complete Count,Tile Count,Shape Count,Pithos Count,Plain Storage Count,Cooking Count,Plain Open Count,Fine Count,Glazed Open Count,Glazed Closed Count,Unknown Type Count,Other Ceramic Type,Other Ceramic Type,Other Type Description" );
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand com = new SqlCommand(sql, connection);
                    connection.Open();
                    SqlDataReader reader = com.ExecuteReader();
                    while (reader.Read())
                    {
                        string unit = reader["Unit"].ToString();
                        string inventory = reader["Inventory_no"].ToString();
                        string bdate = reader["Begin_date"].ToString();
                        string edate = reader["End_date"].ToString();
                        string rimc = reader["Rim_Count"].ToString();
                        string bodyc = reader["Body_Count"].ToString();
                        string handlec = reader["Handle_Count"].ToString();
                        string basec = reader["Base_Count"].ToString();
                        string lidc = reader["Lid_Count"].ToString();
                        string spoutc = reader["Spout_Count"].ToString();
                        string compc = reader["Complete_Count"].ToString();
                        string tilec = reader["Tile_Count"].ToString();
                        string shapec = reader["Shape_Count"].ToString();
                        string pithosc = reader["Pithos_Count"].ToString();
                        string plainc = reader["Plain_Storage_Count"].ToString();
                        string cookc = reader["Cooking_Count"].ToString();
                        string plainoc = reader["Plain_Open_Count"].ToString();
                        string plaincc = reader["Plain_Closed_Count"].ToString();
                        string finec = reader["Fine_Count"].ToString();
                        string glazedcc = reader["Glazed_Cooking_Count"].ToString();
                        string glazedoc = reader["Glazed_Open_Count"].ToString();
                        string glazedclc = reader["Glazed_Closed_Count"].ToString();
                        string unkc = reader["Unknown_Type_Count"].ToString();
                        string otherc = reader["Other_Ceramic_Type"].ToString();
                        string otherd = reader["Other_Type_Description"].ToString();
                        string vesselc = reader["Vessel_Type_Count"].ToString();
                        writer.WriteLine(unit + "," + inventory + "," + bdate + "," + edate + "," + rimc + "," + bodyc + "," + handlec + "," + basec + "," + lidc + "," + spoutc + "," + compc + "," + tilec + "," + shapec + "," + pithosc + "," + plainc + "," + cookc + "," + plainoc + "," + plaincc + "," + finec + "," + glazedcc + "," + glazedoc + "," + glazedclc + "," + unkc + "," + otherc + "," + otherd);
                    }
                }
                writer.Close();
            }
            Response.ContentType = "application/csv";
            Response.AppendHeader("Content-Disposition", "attachment; filename=QueryResult.csv");
            string FilePath = FILE_NAME;
            Response.TransmitFile(FilePath);
            Response.End();
            
        }
    }



}