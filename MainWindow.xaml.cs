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

namespace WindowActivationAndDeactivation
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // 빌드 명령을 수행할 cmd 프로세스 준비
        Process cmdProcess = new Process();
        ProcessStartInfo cmdStartInfo = new ProcessStartInfo();
        
        string msbuildPath = @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe";
        string projectPath = @"C:\Users\W24880\Desktop\Custom_Csharp_Compiler\";
        string projectName = "WPF_IDE.csproj";

        public MainWindow()
        {
            InitializeComponent();

            // RichTextBox 라는 같은 이름의 컴포넌트가 System.Windows.Forms 에 있는데, 얘는 윈폼 컨트롤러고,
            // 우리가 xaml 에서 만든 RichTextBox 는 WPF 의 컨트롤러다. (이름은 같은데 다르다. 속성도 다르게 가짐.)
            FlowDocument myFlowDoc = new FlowDocument();
            rtb_code.Document = myFlowDoc;

            EditingCommands.ToggleInsert.Execute(null, rtb_code);
            
            rtb_code.AppendText("hihi");

            cmdStartInfo.FileName = @"cmd";
            cmdStartInfo.CreateNoWindow = true;
            cmdStartInfo.UseShellExecute = false;
            cmdStartInfo.RedirectStandardOutput = true;
            cmdStartInfo.RedirectStandardInput = true;
            cmdStartInfo.RedirectStandardError = true;
            cmdProcess.StartInfo = cmdStartInfo;
        }

        private void on_Click_Btn_Build(object sender, RoutedEventArgs e) {
            MessageBox.Show("빌드를 시작합니다.");
            try {
                cmdProcess.Start();
                MessageBox.Show("Build 중...");
                cmdProcess.StandardInput.Write(msbuildPath + " " + projectPath + projectName + Environment.NewLine);
                cmdProcess.StandardInput.Close();
                MessageBox.Show(cmdProcess.StandardOutput.ReadToEnd());
                cmdProcess.WaitForExit();
                cmdProcess.Close();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void on_Click_Btn_Run(object sender, RoutedEventArgs e) {
            try {
                cmdProcess.Start();
                MessageBox.Show(projectPath + "WPF_IDE.exe");
                cmdProcess.StandardInput.Write(projectPath + "WPF_IDE.exe");
                cmdProcess.StandardInput.Close();
                MessageBox.Show(cmdProcess.StandardOutput.ReadToEnd());
                cmdProcess.Close();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
    }
}