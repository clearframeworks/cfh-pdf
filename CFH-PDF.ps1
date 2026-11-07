Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$source = @"
using System;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
public sealed class CfhPdfForm : Form {
 readonly TextBox first=new TextBox(),second=new TextBox(),output=new TextBox();
 readonly Label status=new Label(); readonly Button run=new Button(); readonly string engine;
 public CfhPdfForm(string path) {
  engine=path; Text="CFH PDF"; Name="CfhPdf"; StartPosition=FormStartPosition.CenterScreen; ClientSize=new Size(720,330); MinimumSize=new Size(650,355); Font=new Font("Segoe UI",10);
  var title=new Label {Text="Combine two PDFs",Font=new Font("Segoe UI",16,FontStyle.Bold),AutoSize=true,Location=new Point(20,15)};
  var intro=new Label {Text="Pages from the first file come before the second.",AutoSize=true,Location=new Point(22,50)};
  AddInput("First PDF",first,80,false); AddInput("Second PDF",second,145,false); AddInput("Save combined copy as",output,210,true);
  run.Text="Combine"; run.SetBounds(565,270,125,34); run.Anchor=AnchorStyles.Top|AnchorStyles.Right; run.Click+=Combine;
  status.Text="The originals are kept."; status.SetBounds(25,275,525,40); status.Anchor=AnchorStyles.Top|AnchorStyles.Left|AnchorStyles.Right;
  Controls.AddRange(new Control[]{title,intro,run,status}); AcceptButton=run;
 }
 void AddInput(string label,TextBox box,int top,bool save) {
  var l=new Label {Text=label,AutoSize=true,Location=new Point(22,top)}; box.SetBounds(25,top+24,525,28); box.Anchor=AnchorStyles.Top|AnchorStyles.Left|AnchorStyles.Right; box.AccessibleName=label;
  var browse=new Button {Text="Browse...",Location=new Point(565,top+20),Size=new Size(125,34),Anchor=AnchorStyles.Top|AnchorStyles.Right};
  browse.Click+=(s,e)=>{ if(save) { using(var d=new SaveFileDialog {Filter="PDF files|*.pdf",DefaultExt="pdf",AddExtension=true,OverwritePrompt=false,Title="Save combined copy"}) { if(d.ShowDialog(this)==DialogResult.OK) box.Text=d.FileName; } } else { using(var d=new OpenFileDialog {Filter="PDF files|*.pdf",CheckFileExists=true,Title=label}) { if(d.ShowDialog(this)==DialogResult.OK) box.Text=d.FileName; } } };
  Controls.AddRange(new Control[]{l,box,browse});
 }
 static string Quote(string s) { var b=new StringBuilder("\""); int n=0; foreach(char c in s) { if(c=='\\'){n++;continue;} if(c=='\"'){b.Append('\\',n*2+1);b.Append('\"');n=0;continue;} b.Append('\\',n);n=0;b.Append(c); } b.Append('\\',n*2);b.Append('\"');return b.ToString(); }
 async void Combine(object sender,EventArgs e) {
  string a,b,dst;
  try { a=Path.GetFullPath(first.Text); b=Path.GetFullPath(second.Text); dst=Path.GetFullPath(output.Text); } catch {status.Text="Choose valid file paths.";return;}
  if(!File.Exists(a)||!File.Exists(b)){status.Text="Choose two existing PDF files.";return;}
  if(!String.Equals(Path.GetExtension(a),".pdf",StringComparison.OrdinalIgnoreCase)||!String.Equals(Path.GetExtension(b),".pdf",StringComparison.OrdinalIgnoreCase)||!String.Equals(Path.GetExtension(dst),".pdf",StringComparison.OrdinalIgnoreCase)){status.Text="Choose PDF files and a PDF destination.";return;}
  if(String.Equals(a,dst,StringComparison.OrdinalIgnoreCase)||String.Equals(b,dst,StringComparison.OrdinalIgnoreCase)){status.Text="Choose a different name for the copy.";return;}
  if(File.Exists(dst)){status.Text="That copy already exists. Choose another name.";return;}
  if(!File.Exists(engine)){status.Text="PDF engine missing. Reinstall CFH PDF.";return;}
  string temp=Path.Combine(Path.GetDirectoryName(dst),".cfh-"+Guid.NewGuid().ToString("N")+".pdf"); run.Enabled=false; status.Text="Combining...";
  try {
   var p=new Process {StartInfo=new ProcessStartInfo {FileName=engine,Arguments="--empty --pages "+Quote(a)+" "+Quote(b)+" -- "+Quote(temp),UseShellExecute=false,CreateNoWindow=true,RedirectStandardError=true,RedirectStandardOutput=true}};
   p.Start(); var err=p.StandardError.ReadToEndAsync(); var stdout=p.StandardOutput.ReadToEndAsync(); await Task.Run(()=>p.WaitForExit()); await Task.WhenAll(err,stdout);
   if(p.ExitCode!=0||!File.Exists(temp)){status.Text="Could not combine these PDFs. Check the files and try again.";}
   else {File.Move(temp,dst);status.Text="Combined copy saved.";}
  } catch(Exception) {status.Text="Could not save the copy. Check the folder and try again.";}
  finally {try{if(File.Exists(temp))File.Delete(temp);}catch{}run.Enabled=true;}
 }
}
"@
Add-Type -TypeDefinition $source -ReferencedAssemblies System.Windows.Forms,System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run((New-Object CfhPdfForm (Join-Path $PSScriptRoot 'qpdf.exe')))