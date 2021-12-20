using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.Reflection;

using Excel = Microsoft.Office.Interop.Excel;

namespace WindowActivationAndDeactivation
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            EmployeeList employeeList = new EmployeeList();

            Employee employee1 = new Employee();

            employee1.ID          = 52590;
            employee1.Name        = "Sathish";
            employee1.Designation = "Developer";            

            employeeList.Add(employee1);

            Employee employee2 = new Employee();

            employee2.ID          = 52592;
            employee2.Name        = "Karthick";
            employee2.Designation = "Developer";            

            employeeList.Add(employee2);

            Employee employee3 = new Employee();

            employee3.ID          = 52593;
            employee3.Name        = "Raja";
            employee3.Designation = "Manager";            

            employeeList.Add(employee3);

            Employee employee4 = new Employee();

            employee4.ID          = 12778;
            employee4.Name        = "Sumesh";
            employee4.Designation = "Project Lead";            

            employeeList.Add(employee4);

            Employee employee5 = new Employee();

            employee5.ID          = 12590;
            employee5.Name        = "Srini";
            employee5.Designation = "Project Lead";            

            employeeList.Add(employee5);

            this.employeeDataGrid.ItemsSource = employeeList;
        }

        private void exportEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            ExcelExportHelper<Employee, EmployeeList> helper = new ExcelExportHelper<Employee, EmployeeList>();

            helper.List = employeeDataGrid.ItemsSource as EmployeeList;

            helper.Generate();
        }
    }

    public class EmployeeList : List<Employee>
    {
    }

    public class Employee
    {
        //////////////////////////////////////////////////////////////////////////////////////////////////// Property
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region ID - ID

        /// <summary>
        /// ID
        /// </summary>
        public int ID { get; set; }

        #endregion
        #region 성명 - Name

        /// <summary>
        /// 성명
        /// </summary>
        public string Name { get; set; }

        #endregion
        #region 직위 - Designation

        /// <summary>
        /// 직위
        /// </summary>
        public string Designation { get; set; }

        #endregion
    }

    public class ExcelExportHelper<TEntity, TList> where TEntity : class where TList : List<TEntity>
    {
        //////////////////////////////////////////////////////////////////////////////////////////////////// Field
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region Field

        /// <summary>
        /// 리스트
        /// </summary>
        public List<TEntity> List;

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////// Private

        #region Field

        /// <summary>
        /// 응용 프로그램
        /// </summary>
        private Microsoft.Office.Interop.Excel.Application application = null;

        /// <summary>
        /// 워크북들
        /// </summary>
        private Microsoft.Office.Interop.Excel.Workbooks workbooks = null;

        /// <summary>
        /// 워크북
        /// </summary>
        private Microsoft.Office.Interop.Excel._Workbook workbook = null;

        /// <summary>
        /// 시트들
        /// </summary>
        private Microsoft.Office.Interop.Excel.Sheets sheets = null;

        /// <summary>
        /// 워크시트
        /// </summary>
        Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

        /// <summary>
        /// 범위
        /// </summary>
        private Microsoft.Office.Interop.Excel.Range range = null;

        /// <summary>
        /// 폰트
        /// </summary>
        private Microsoft.Office.Interop.Excel.Font font = null;

        /// <summary>
        /// 부가적인 값
        /// </summary>
        private object optionalValue = Missing.Value;

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////////// Method
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region 생성하기 - Generate()

        /// <summary>
        /// 생성하기
        /// </summary>
        public void Generate()
        {
            try
            {
                if(List != null)
                {
                    if(List.Count != 0)
                    {
                        CreateInstance();

                        FillSheet();

                        ShowEXCEL();
                    }
                }
            }
            finally
            {
                ReleaseInstance(this.worksheet  );
                ReleaseInstance(this.sheets     );
                ReleaseInstance(this.workbook   );
                ReleaseInstance(this.workbooks  );
                ReleaseInstance(this.application);
            }
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////// Private

        #region 인스턴스 생성하기 - CreateInstance()

        /// <summary>
        /// 인스턴스 생성하기
        /// </summary>
        private void CreateInstance()
        {
            this.application = new Microsoft.Office.Interop.Excel.Application();

            this.workbooks = (Microsoft.Office.Interop.Excel.Workbooks)this.application.Workbooks;

            this.workbook = (Microsoft.Office.Interop.Excel._Workbook)(this.workbooks.Add(this.optionalValue));

            this.sheets = (Microsoft.Office.Interop.Excel.Sheets)this.workbook.Worksheets;

            this.worksheet = this.sheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
        }

        #endregion
        #region 행들 추가하기 - AddRows(startRange, rowCount, columnCount, values)

        /// <summary>
        /// 행들 추가하기
        /// </summary>
        /// <param name="startRange">시작 범위</param>
        /// <param name="rowCount">행 수</param>
        /// <param name="columnCount">열 수</param>
        /// <param name="values">값들</param>
        private void AddRows(string startRange, int rowCount, int columnCount, object values)
        {
            this.range = this.worksheet.get_Range(startRange, this.optionalValue);

            this.range = this.range.get_Resize(rowCount, columnCount);

            this.range.set_Value(this.optionalValue, values);
        }       

        #endregion
        #region 헤더 스타일 설정하기 - SetHeaderStyle()

        /// <summary>
        /// 헤더 스타일 설정하기
        /// </summary>
        private void SetHeaderStyle()
        {
            this.font = this.range.Font;

            this.font.Bold = true;
        }

        #endregion
        #region 헤더 생성하기 - CreateHeader()

        /// <summary>
        /// 헤더 생성하기
        /// </summary>
        /// <returns>헤더 목록</returns>
        private object[] CreateHeader()
        {
            PropertyInfo[] propertyInfoArray = typeof(TEntity).GetProperties();

            List<object> headerList = new List<object>();

            for(int i = 0; i < propertyInfoArray.Length; i++)
            {
                headerList.Add(propertyInfoArray[i].Name);
            }

            object[] headerArray = headerList.ToArray();

            AddRows("A1", 1, headerArray.Length, headerArray);

            SetHeaderStyle();

            return headerArray;
        }

        #endregion
        #region 컬럼 너비 자동 맞추기 - AutoFitColumnWidth(startRange, rowCount, columnCount)

        /// <summary>
        /// 컬럼 너비 자동 맞추기
        /// </summary>
        /// <param name="startRange">시작 범위</param>
        /// <param name="rowCount">행 수</param>
        /// <param name="columnCount">열 수</param>
        private void AutoFitColumnWidth(string startRange, int rowCount, int columnCount)
        {
            this.range = this.worksheet.get_Range(startRange, this.optionalValue);

            this.range = this.range.get_Resize(rowCount, columnCount);

            this.range.Columns.AutoFit();
        }

        #endregion
        #region 데이타 작성하기 - WriteData(headerArray)

        /// <summary>
        /// 데이타 작성하기
        /// </summary>
        /// <param name="headerArray">헤더 배열</param>
        private void WriteData(object[] headerArray)
        {
            object[,] dataArray = new object[List.Count, headerArray.Length];

            for(int row = 0; row < List.Count; row++)
            {
                TEntity entity = List[row];

                for(int column = 0; column < headerArray.Length; column++)
                {
                    object value = typeof(TEntity).InvokeMember
                    (
                        headerArray[column].ToString(),
                        BindingFlags.GetProperty,
                        null,
                        entity,
                        null
                    );

                    dataArray[row, column] = (value == null) ? "" : value.ToString();
                }
            }

            AddRows("A2", List.Count, headerArray.Length, dataArray);

            AutoFitColumnWidth("A1", List.Count + 1, headerArray.Length);
        }

        #endregion
        #region 시트 채우기 - FillSheet()

        /// <summary>
        /// 시트 채우기
        /// </summary>
        private void FillSheet()
        {
            object[] headerArray = CreateHeader();

            WriteData(headerArray);
        }

        #endregion
        #region 엑셀 보여주기 - ShowEXCEL()

        /// <summary>
        /// 엑셀 보여주기
        /// </summary>
        private void ShowEXCEL()
        {
            this.application.Visible = true;
        }

        #endregion
        #region 인스턴스 해제하기 - ReleaseInstance(instance)

        /// <summary>
        /// 인스턴스 해제하기
        /// </summary>
        /// <param name="instance">인스턴스</param>
        private void ReleaseInstance(object instance)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(instance);

                instance = null;
            }
            catch(Exception exception)
            {
                instance = null;

                throw exception;
            }
            finally
            {
                GC.Collect();
            }
        }

        #endregion
    }
}