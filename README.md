# Line_senstive_Frist_step_Ini
click button. show line chart. WPF,
  public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private List<string> strListx = new List<string>() {"1","20","40","80","100","120","140","160","180","200","220","240" };
       private List<string> strListy = new List<string>() ;
       // private List<double> strListy = new List<double>();
        private List<string> strListz = new List<string>() { "54", "65", "34", "28", "76", "72", "19", "34" };

        private void button_run_Click(object sender, RoutedEventArgs e)
        {
            //String pathstring = "C:\\Users\\z003rx8n.AD001\\Desktop\\Line_Sensitive\\Line_Sensitive001\\bin\\Debug\\update_line distances.xlsx";
            string str4 = AppDomain.CurrentDomain.BaseDirectory;

             String pathstring = Path.Combine(str4, @"update_line distances.xlsx");
            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathstring + ";Extended Properties=Excel 12.0";
            con.Open();
            string strCom = " SELECT * FROM ["+CableTypeText.Text+"$] ";
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, con);
            DataSet myDataSet = new DataSet();

            myCommand.Fill(myDataSet, "["+CableTypeText.Text+"$]");
            for (int i = 0; i < myDataSet.Tables[0].Rows.Count; i++)
            {
              //  if (myDataSet.Tables[0].Rows[i]["Distance[kM]"] != null)
                //    strListx[i]= myDataSet.Tables[0].Rows[i]["Distance[kM]"].ToString();
                if (myDataSet.Tables[0].Rows[i][""+VoltageText.Text+"kV"] != null)
                    strListy.Add(myDataSet.Tables[0].Rows[i][""+VoltageText.Text +"kV"].ToString());
                   // strListy[i] = myDataSet.Tables[0].Rows[i]["150kV"].ToString();
             // getst= myDataSet.Tables[0].Rows[i]["150kV"].ToString();
             //   strListy[i] = getst;
               // tb1.Text += myDataSet.Tables[0].Rows[i]["150kV"].ToString();
               // tb1.Text += getst + "\n";
            }
            con.Close();


           
             Chart chart = new Chart();

           
            chart.Width = 580;
            chart.Height = 380;
            chart.Margin = new Thickness(0, 0, 0, 0);
          

            chart.ToolBarEnabled = false;

          
            chart.ScrollingEnabled = true ;
            chart.View3D = true;

          
            Title title = new Title();

           
            title.Text = "Active Power Transmission capability ";
            title.Padding = new Thickness(0, 10, 5, 0);

            //向图标添加标题
            chart.Titles.Add(title);

            Axis yAxis = new Axis();
                     
            yAxis.AxisMinimum = 0;
                     
            yAxis.Suffix = " [kV]";
           
            chart.AxesY.Add(yAxis);

                        
            DataSeries dataSeries = new DataSeries();

            dataSeries.RenderAs = RenderAs.Spline;//柱状Stacked

        
            DataPoint dataPoint;
            for (int i = 0; i < strListx.Count; i++)
            {
                                   
                dataPoint = new DataPoint();
                                  
                dataPoint.AxisXLabel = strListx[i];
                                  
                dataPoint.YValue = double.Parse(strListy[i]);
                       
               // dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                                  
                dataSeries.DataPoints.Add(dataPoint);
            }

                       
            chart.Series.Add(dataSeries);

            Chart chart1 = new Chart();
            
            gr1.Children.Add(chart);

            p2.Children.Add(chart1);
          
        }
           
        private void tb1_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Slider_T_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            List<string> Tstr = new List<string>() {"AlSt45040", "AlSt38050", "AlSt24040", "AlSt15025"};
           
            int x=0;
           
            x = Convert.ToInt32(Slider_T.Value);
            CableTypeText.Text=Tstr[x];
            
        }

        private void Slider_V_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

   
        }


    }
}
