using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;
namespace AppProgramAdder;
internal static class Program
{
    [STAThread] static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        try
        {
            if(args.Length==2 && args[0]=="--add")
            {
                using var dialog=new AddDialog(args[1]);
                dialog.ShowDialog();
                return;
            }
            if(args.Length==0)Application.Run(new MainForm());
            else MessageBox.Show("Open the app normally or use Add to App Program Adder on a file.","App Program Adder");
        }
        catch(Exception ex){MessageBox.Show(ex.Message,"App Program Adder",MessageBoxButtons.OK,MessageBoxIcon.Error);}
    }
}
internal static class Tools
{
    internal static readonly string Programs=Environment.GetFolderPath(Environment.SpecialFolder.Programs);
    internal static void SetIcon(Form form)
    {
        try{string? exe=Environment.ProcessPath;if(exe!=null)form.Icon=Icon.ExtractAssociatedIcon(exe);}catch{}
    }
    internal static string Add(string target,string label)
    {
        target=Path.GetFullPath(target);label=label.Trim();
        if(!File.Exists(target))throw new IOException("Choose an existing file first.");
        if(string.IsNullOrWhiteSpace(label)||label.Length>100||label.IndexOfAny(Path.GetInvalidFileNameChars())>=0||label.EndsWith('.'))
            throw new ArgumentException("Enter a name of 1–100 characters, without filename symbols or a trailing period.");
        string stem=label.Split('.')[0].TrimEnd();
        if(new[]{"CON","PRN","AUX","NUL","CONIN$","CONOUT$"}.Contains(stem,StringComparer.OrdinalIgnoreCase)||
            (stem.Length==4&&(stem.StartsWith("COM",StringComparison.OrdinalIgnoreCase)||stem.StartsWith("LPT",StringComparison.OrdinalIgnoreCase))&&"123456789¹²³".Contains(stem[3])))
            throw new ArgumentException("That name is reserved by Windows. Choose another.");
        Directory.CreateDirectory(Programs);string destination=Path.Combine(Programs,label+".lnk");
        if(File.Exists(destination)||Directory.Exists(destination))throw new IOException("That shortcut name already exists. Nothing was overwritten.");
        string temp=Path.Combine(Programs,"AppProgramAdder-"+Guid.NewGuid().ToString("N")+".lnk");
        object? shell=null,link=null;
        try
        {
            if(Path.GetExtension(target).Equals(".lnk",StringComparison.OrdinalIgnoreCase))File.Copy(target,temp,false);
            else
            {
                var type=Type.GetTypeFromProgID("WScript.Shell")??throw new InvalidOperationException("Windows Script Host is unavailable.");
                shell=Activator.CreateInstance(type)??throw new InvalidOperationException("Cannot create the Windows Shell object.");
                link=((dynamic)shell).CreateShortcut(temp);dynamic shortcut=link;
                shortcut.TargetPath=target;shortcut.WorkingDirectory=Path.GetDirectoryName(target)??"";
                shortcut.Description="Added with App Program Adder";
                if(Path.GetExtension(target).Equals(".exe",StringComparison.OrdinalIgnoreCase))shortcut.IconLocation=target+",0";
                shortcut.Save();
            }
            File.Move(temp,destination,false);return destination;
        }
        finally
        {
            if(link!=null&&Marshal.IsComObject(link))Marshal.FinalReleaseComObject(link);
            if(shell!=null&&Marshal.IsComObject(shell))Marshal.FinalReleaseComObject(shell);
            try{if(File.Exists(temp))File.Delete(temp);}catch{}
        }
    }
}
internal static class ExplorerEntry
{
    const string KeyPath=@"Software\Classes\*\shell\AppProgramAdder.AddToStart";
    internal static bool Enabled(){using var k=Registry.CurrentUser.OpenSubKey(KeyPath);return k!=null;}
    internal static bool CurrentLocation()
    {
        using var k=Registry.CurrentUser.OpenSubKey(KeyPath+@"\command");
        return string.Equals(k?.GetValue("") as string,Command(),StringComparison.OrdinalIgnoreCase);
    }
    static string Exe()=>Environment.ProcessPath??throw new InvalidOperationException("Cannot find this app's EXE.");
    static string Command()=>"\""+Exe()+"\" --add \"%1\"";
    internal static void Enable()
    {
        string exe=Exe();
        if(!Path.GetFileName(exe).Equals("AppProgramAdder.exe",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Build and open ReadyToRun\\AppProgramAdder.exe before enabling the entry.");
        using(var k=Registry.CurrentUser.CreateSubKey(KeyPath))
        {
            k.SetValue("","Add to App Program Adder");k.SetValue("Icon","\""+exe+"\",0");k.SetValue("MultiSelectModel","Single");
            using var command=k.CreateSubKey("command");command.SetValue("",Command());
        }
        Notify();
    }
    internal static void Remove(){Registry.CurrentUser.DeleteSubKeyTree(KeyPath,false);Notify();}
    static void Notify()=>SHChangeNotify(0x08000000,0,IntPtr.Zero,IntPtr.Zero);
    [DllImport("shell32.dll")]static extern void SHChangeNotify(uint id,uint flags,IntPtr item1,IntPtr item2);
}
public sealed class AddDialog:Form
{
    readonly string target;
    readonly TextBox name=new(){Dock=DockStyle.Fill,MaxLength=100};
    readonly Button ok=new(){Text="OK",AutoSize=true,MinimumSize=new Size(85,30)};
    bool saving;
    public AddDialog(string path)
    {
        target=Path.GetFullPath(path);Text="Add to App Program Adder";Tools.SetIcon(this);
        Font=new Font("Segoe UI",10);AutoScaleMode=AutoScaleMode.Dpi;StartPosition=FormStartPosition.CenterScreen;
        FormBorderStyle=FormBorderStyle.FixedDialog;MaximizeBox=false;MinimizeBox=false;
        AutoSize=true;AutoSizeMode=AutoSizeMode.GrowAndShrink;MinimumSize=new Size(510,230);
        var layout=new TableLayoutPanel{AutoSize=true,Dock=DockStyle.Fill,Padding=new Padding(18),ColumnCount=1,RowCount=4};
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,470));
        var info=new Label{Text="Create a Start menu shortcut to:\n"+target,AutoSize=true,MaximumSize=new Size(470,100),Margin=new Padding(3,3,3,12)};
        layout.Controls.Add(info,0,0);layout.Controls.Add(new Label{Text="Shortcut name",AutoSize=true},0,1);
        name.Text=Path.GetFileNameWithoutExtension(target);layout.Controls.Add(name,0,2);
        var buttons=new FlowLayoutPanel{AutoSize=true,Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft,Margin=new Padding(0,12,0,0)};
        var cancel=new Button{Text="Cancel",AutoSize=true,MinimumSize=new Size(85,30)};
        cancel.Click+=(_,_)=>{if(saving)return;DialogResult=DialogResult.Cancel;Close();};
        ok.Click+=(_,_)=>Save();buttons.Controls.Add(cancel);buttons.Controls.Add(ok);layout.Controls.Add(buttons,0,3);
        AcceptButton=ok;CancelButton=cancel;Controls.Add(layout);
        Shown+=(_,_)=>{name.Focus();name.SelectAll();};
    }
    void Save()
    {
        if(saving)return;saving=true;ok.Enabled=false;
        try{Tools.Add(target,name.Text);DialogResult=DialogResult.OK;Close();}
        catch(Exception ex){MessageBox.Show(this,ex.Message,"Could not add shortcut",MessageBoxButtons.OK,MessageBoxIcon.Warning);name.Focus();}
        finally{saving=false;if(!IsDisposed)ok.Enabled=true;}
    }
}
public sealed class MainForm:Form
{
    readonly TextBox file=new(){Dock=DockStyle.Fill,ReadOnly=true},name=new(){Dock=DockStyle.Fill,MaxLength=100};
    readonly Label status=new(){AutoSize=true,MaximumSize=new Size(610,0),Text="Choose any file. Adding it will not run it."};
    readonly Label menuStatus=new(){AutoSize=true,MaximumSize=new Size(610,0)};
    readonly Button toggle=new(){AutoSize=true};
    public MainForm()
    {
        Text="App Program Adder";ClientSize=new Size(690,480);MinimumSize=new Size(700,510);StartPosition=FormStartPosition.CenterScreen;
        Font=new Font("Segoe UI",10);BackColor=Color.FromArgb(245,247,251);Tools.SetIcon(this);
        var layout=new TableLayoutPanel{Dock=DockStyle.Fill,Padding=new Padding(22),ColumnCount=2,RowCount=10,AutoScroll=true};
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,115));
        void Wide(Control c,int row){layout.Controls.Add(c,0,row);layout.SetColumnSpan(c,2);}
        Wide(new Label{Text="App Program Adder",AutoSize=true,Font=new Font("Segoe UI",19,FontStyle.Bold)},0);
        Wide(new Label{Text="Add any file shortcut to your Start menu. The original stays put.",AutoSize=true,Margin=new Padding(3,10,3,15)},1);
        layout.Controls.Add(file,0,2);var browse=new Button{Text="Browse...",Dock=DockStyle.Fill,AutoSize=true};browse.Click+=(_,_)=>Browse();layout.Controls.Add(browse,1,2);
        Wide(new Label{Text="Shortcut name",AutoSize=true,Margin=new Padding(3,15,3,5)},3);Wide(name,4);
        var buttons=new FlowLayoutPanel{AutoSize=true,Dock=DockStyle.Fill,Margin=new Padding(0,15,0,12)};
        var add=new Button{Text="Add to Start menu",AutoSize=true};add.Click+=(_,_)=>Add();buttons.Controls.Add(add);
        var open=new Button{Text="Open Programs folder",AutoSize=true};open.Click+=(_,_)=>{
            try{Directory.CreateDirectory(Tools.Programs);var si=new ProcessStartInfo("explorer.exe"){UseShellExecute=true};si.ArgumentList.Add(Tools.Programs);Process.Start(si);}
            catch(Exception ex){MessageBox.Show(this,ex.Message);}};
        buttons.Controls.Add(open);Wide(buttons,5);Wide(status,6);
        Wide(new Label{Text="File Explorer right-click menu",AutoSize=true,Font=new Font("Segoe UI",11,FontStyle.Bold),Margin=new Padding(3,22,3,5)},7);
        var action=new FlowLayoutPanel{AutoSize=true,Dock=DockStyle.Fill};action.Controls.Add(toggle);Wide(action,8);Wide(menuStatus,9);
        toggle.Click+=(_,_)=>ToggleMenu();Controls.Add(layout);Activated+=(_,_)=>RefreshMenu();RefreshMenu();
    }
    void Browse()
    {
        using var picker=new OpenFileDialog{Title="Choose any program or file",Filter="All files (*.*)|*.*",CheckFileExists=true,DereferenceLinks=false};
        if(picker.ShowDialog(this)!=DialogResult.OK)return;file.Text=picker.FileName;name.Text=Path.GetFileNameWithoutExtension(file.Text);status.Text="Ready. Adding a shortcut will not run this file.";
    }
    void Add()
    {
        if(!File.Exists(file.Text)){MessageBox.Show(this,"Choose an existing file first.");return;}
        if(MessageBox.Show(this,$"Create this Start menu shortcut?\n\nName: {name.Text.Trim()}\n\nFile: {file.Text}","Confirm shortcut",MessageBoxButtons.OKCancel,MessageBoxIcon.Question)!=DialogResult.OK)return;
        try{string destination=Tools.Add(file.Text,name.Text);status.Text="Shortcut created: "+Path.GetFileNameWithoutExtension(destination)+". Open Start to find it.";}
        catch(Exception ex){MessageBox.Show(this,ex.Message,"Could not add shortcut",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
    }
    void RefreshMenu()
    {
        try
        {
            bool enabled=ExplorerEntry.Enabled();toggle.Enabled=true;toggle.Text=enabled?"Remove right-click entry":"Enable right-click entry";
            menuStatus.Text=enabled?(ExplorerEntry.CurrentLocation()?"Enabled. On Windows 11, check Show more options if needed.":"Enabled for a different location. Remove, then enable again to use this copy."):"Not enabled. Add the entry to your file right-click menu.";
        }
        catch{toggle.Enabled=false;menuStatus.Text="Could not read the right-click setting.";}
    }
    void ToggleMenu()
    {
        try
        {
            bool enabled=ExplorerEntry.Enabled();
            string question=enabled?"Remove this app's right-click entry? Existing Start menu shortcuts will stay.":"Add 'Add to App Program Adder' to your right-click menu?\n\nThis changes only your Windows account's menu settings and uses this EXE's current location. It will not change file associations or run selected files.";
            if(MessageBox.Show(this,question,"App Program Adder",MessageBoxButtons.OKCancel,MessageBoxIcon.Question)!=DialogResult.OK)return;
            if(enabled)ExplorerEntry.Remove();else ExplorerEntry.Enable();
        }
        catch(Exception ex){MessageBox.Show(this,ex.Message,"Could not change menu",MessageBoxButtons.OK,MessageBoxIcon.Error);}
        finally{RefreshMenu();}
    }
}
