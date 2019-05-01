using System;
using System.Data;
using System.Data.Odbc;
using System.Linq;

namespace TOMSSQL
{
    class ExcelParser
    {
        public void CreateTablesInMSSQL(OdbcConnection connection)
        {
            string sql = @"
                        use example
                        go

                        if OBJECT_ID('dbo.Clubs','U') is not null
	                        drop table dbo.Clubs
                        go

                        create table dbo.Clubs 
                        (clubID int not null identity,
                        clubName nvarchar(30),
                        clubCity  nvarchar(30) not null,

                        constraint PK_Club primary key(clubID)
                        )
                        go

                        use example
                        go

                        if OBJECT_ID('dbo.Swimmers','U') is not null
	                        drop table dbo.Swimmers
                        go

                        create table dbo.Swimmers 
                        (
                        swimmerID int not null identity,
                        lastName nvarchar(30) not null,
                        firstName  nvarchar(30) not null,
                        yearOfBirth int not null,
                        clubID int,
                        constraint PK_Swimmer  primary key(swimmerID)

                        )
                        go

                        use example
                        go

                        if OBJECT_ID('dbo.Disciplines','U') is not null
	                        drop table dbo.Disciplines
                        go

                        create table dbo.Disciplines 
                        (
                        disciplineID int not null identity,
                        distance int not null,
                        gender nvarchar(10),
                        style nvarchar(20) not null,
                        compYear int not null,

                        constraint PK_Discipline  primary key(disciplineID)
                        )
                        go

                        use example
                        go

                        if OBJECT_ID('dbo.Results','U') is not null
	                        drop table dbo.Results
                        go

                        create table dbo.Results 
                        (
                        resultID int not null identity,
                        place int not null,
                        resultTime nvarchar(20) not null,
                        swimmerID int,
                        disciplineID int,
                        constraint PK_Result primary key(resultID)
                        )
                        go

                        alter table dbo.Swimmers 
	                        add constraint FK_Swimmer_Club foreign key (clubID) references dbo.Clubs(clubID)
                        go

                        alter table dbo.Results 
	                        add constraint FK_Result_Swimmer foreign key (swimmerID) references dbo.Swimmers(swimmerID)
                        go

                        alter table dbo.Results 
	                        add constraint FK_Result_Discipline foreign key (disciplineID) references dbo.Disciplines(disciplineID)
                        go";

            OdbcCommand command = new OdbcCommand(sql, connection);
            command.ExecuteNonQuery();
        }

        private string GetExcelSheetName(OdbcConnection conn, int index)
        {

            DataTable table = conn.GetSchema("Tables");
            string sheetName = table.Rows[index][2].ToString();
            return sheetName;
        }

        public DataSet GetDataSetFromExcel()
        {
            string FilePath = "..\\..\\resources\\Competition2.xls";
            string odbcConnetionString = "Driver={Microsoft Excel Driver (*.xls)};DBQ=" + FilePath + ";";

            using (OdbcConnection conn = new OdbcConnection(odbcConnetionString))
            {
                conn.Open();
                string sheetName = GetExcelSheetName(conn, 1);
                OdbcDataAdapter odbcDA = new OdbcDataAdapter("select * from [" + sheetName + "]", conn);
                DataSet excelDataSet = new DataSet();
                odbcDA.Fill(excelDataSet);
                return excelDataSet;
            }
        }

        public OdbcConnection CreateConnectionToMSSQL()
        {

            string odbcConnetionString = "Driver={SQL Server};Server=10.1.1.85;Database=EXAMPLE;Trusted_Connection = Yes;";

            return new OdbcConnection(odbcConnetionString);

        }


        public void ProcessRows(OdbcConnection conn, DataRowCollection rows)
        {
            int compYear;
            int distance;
            string style = "";
            string gender = "";
            int disciplineId = 0;

            foreach (DataRow row in rows)
            {
                int place;
                string points;

                string lastName;
                string firstName;
                string clubCity;
                string clubName;
                string resultTime;
                int yearOfBirth;


                object[] arrayWithData = row.ItemArray;

                if (arrayWithData[0].ToString() != "")
                {
                    place = Convert.ToInt32(arrayWithData[0].ToString());
                    lastName = arrayWithData[1].ToString().Split(' ')[0];
                    firstName = arrayWithData[1].ToString().Split(' ')[1];
                    yearOfBirth = Convert.ToInt32(arrayWithData[2].ToString());

                    if (arrayWithData[4].ToString().Split(',').Length == 2)
                    {
                        clubCity = arrayWithData[4].ToString().Split(',')[0];
                        clubName = arrayWithData[4].ToString().Split(',')[1];
                    }
                    else
                    {
                        clubCity = arrayWithData[4].ToString().Split(',')[0];
                        clubName = "";
                    }
                    resultTime = arrayWithData[5].ToString();
                    //points = arrayWithData[8].ToString();

                    int clubId = InsertRowIntoTableClubs(conn, clubCity, clubName);
                    int swimmerId = InsertRowIntoSwimmers(conn, lastName, firstName, yearOfBirth, clubId);
                    InsertRowIntoResults(conn, swimmerId, disciplineId, place, resultTime);
                }
                else if (arrayWithData[1].ToString() != "")
                {
                    string[] arr = arrayWithData[1].ToString().Split(' ');
                    arr = arr.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                    if (arr.Length < 4)
                    {
                        continue;
                    }
                    distance = Convert.ToInt32(arr[0]);
                    style = arr[1];
                    gender = arr[2];
                    compYear = Convert.ToInt32(arr[3]);

                    disciplineId = InsertRowIntoDisciplines(conn, distance, style, gender, compYear);
                }

            }
        }


        private int InsertRowIntoTableClubs(OdbcConnection connection, string city, string name)
        {
            string sql = String.Format("select clubID from Clubs where clubCity='{0}' and clubName='{1}'", city, name);
            OdbcCommand command = new OdbcCommand(sql, connection);

            object clubID = command.ExecuteScalar();

            if (clubID == null)
            {
                sql = String.Format("insert into Clubs(clubCity,clubName) values('{0}','{1}'); select @@IDENTITY", city, name);
                command.CommandText = sql;
                clubID = command.ExecuteScalar();
            }
            return Convert.ToInt32(clubID);
            //int clubId = Int32.Parse(command.ExecuteScalar().ToString());
        }
        private int InsertRowIntoDisciplines(OdbcConnection connection, int distance, string style, string gender, int compYear)
        {
            string sql = String.Format("select disciplineID from Disciplines where distance={0} and style='{1}' and gender='{2}' and compYear={3}", distance, style, gender, compYear);
            OdbcCommand command = new OdbcCommand(sql, connection);

            object disciplineID = command.ExecuteScalar();

            if (disciplineID == null)
            {
                sql = String.Format("insert into Disciplines(distance, style, gender, compYear) values({0},'{1}','{2}',{3}) select @@IDENTITY", distance, style, gender, compYear);
                command.CommandText = sql;
                disciplineID = command.ExecuteScalar();
            }
            return Convert.ToInt32(disciplineID);

        }

        private void InsertRowIntoResults(OdbcConnection connection, int swimmerID, int disciplineID, int place, string resultTime)
        {

            string sql = String.Format("insert into Results(swimmerID, disciplineID, place, resultTime) values({0},{1},{2},'{3}')", swimmerID, disciplineID, place, resultTime);

            OdbcCommand command = new OdbcCommand(sql, connection);
            command.ExecuteNonQuery();

        }

        private int InsertRowIntoSwimmers(OdbcConnection connection, string lastName, string firstName, int yearOfBirth, int clubID)
        {
            string sql = String.Format("select swimmerID from Swimmers where lastName='{0}' and firstName='{1}' and yearOfBirth='{2}' and clubID={3}", lastName, firstName, yearOfBirth, clubID);
            OdbcCommand command = new OdbcCommand(sql, connection);

            object swimmerID = command.ExecuteScalar();

            if (swimmerID == null)
            {
                sql = String.Format("insert into Swimmers(lastName,firstName,yearOfBirth,clubID) values('{0}','{1}','{2}',{3}) select @@IDENTITY", lastName, firstName, yearOfBirth, clubID);
                command.CommandText = sql;
                swimmerID = command.ExecuteScalar();
            }
            return Convert.ToInt32(swimmerID);

        }



    }
}
