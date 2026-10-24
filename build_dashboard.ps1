$ErrorActionPreference='Stop'
$root=$PSScriptRoot
Add-Type -TypeDefinition @'
using System; using System.IO; using System.Text; using System.Collections.Generic; using System.Globalization;
public class DTBuilder {
  class P { public string code,vmpp,name,pack,form; public P(string a,string b,string c,string d,string e){code=a;vmpp=b;name=c;pack=d;form=e;} }
  class O { public int d,p,c; public double v; public O(int a,int b,double x,int z){d=a;p=b;v=x;c=z;} }
  static string S(object x){return x==null?"":Convert.ToString(x,CultureInfo.InvariantCulture).Trim();}
  static double N(object x){double z; return double.TryParse(S(x),NumberStyles.Any,CultureInfo.InvariantCulture,out z)?z:0;}
  static string J(string s){return "\""+(s??"").Replace("\\","\\\\").Replace("\"","\\\"").Replace("\r","\\r").Replace("\n","\\n").Replace("\t","\\t")+"\"";}
  static void Add(List<object[]> r,string date,object pack,object price,string code,string cat,string vmpp,string name,string form){if(price==null||S(price)=="")return;r.Add(new object[]{date,S(pack),N(price),code,cat,vmpp,name,form});}
  public static void Build(object[,] h,object[,] a,object[,] b,object[,] d,object[,] m,string folder){
    var rows=new List<object[]>(h.GetLength(0)+5000);
    for(int i=2;i<=h.GetLength(0);i++){var date=DateTime.FromOADate(N(h[i,1])).ToString("yyyy-MM-01");var vm=S(h[i,6]);Add(rows,date,h[i,2],h[i,3],S(h[i,4]),S(h[i,5]),vm,vm,"");}
    for(int i=4;i<=a.GetLength(0);i++){var name=S(a[i,1]); if(name!="")Add(rows,"2026-06-01",a[i,2],a[i,5],name,S(a[i,4]),name+" "+S(a[i,2])+" "+S(a[i,3]),name,S(a[i,3]));}
    for(int i=4;i<=b.GetLength(0);i++){var name=S(b[i,1]); if(name!="")Add(rows,"2026-06-01",b[i,2],b[i,6],name,"Part VIIIB",name+" "+S(b[i,2])+" "+S(b[i,3]),name,S(b[i,3]));}
    for(int i=4;i<=d.GetLength(0);i++){var name=S(d[i,1]); if(name!="")Add(rows,"2026-06-01",d[i,2],d[i,6],name,"Part VIIID",name+" "+S(d[i,2])+" "+S(d[i,3]),name,S(d[i,3]));}
    for(int i=4;i<=m.GetLength(0);i++){var name=S(m[i,2]); if(name!="")Add(rows,"2026-06-01",m[i,3],m[i,5],S(m[i,1]),"Part VIIIA Category M",name+" "+S(m[i,3])+" "+S(m[i,4]),name,S(m[i,4]));}
    using(var sw=new StreamWriter(Path.Combine(folder,"Drug_Tariff_Consolidated_to_Jun_2026.csv"),false,new UTF8Encoding(true))){sw.WriteLine("date,pack_size,price_pence,product,tariff_category,vmpp,source_name,dosage_form");foreach(var x in rows)sw.WriteLine(string.Join(",",Array.ConvertAll(x,z=>"\""+S(z).Replace("\"","\"\"")+"\"")));}
    var dates=new List<string>();var dateId=new Dictionary<string,int>();var cats=new List<string>();var catId=new Dictionary<string,int>();var products=new List<P>();var prodId=new Dictionary<string,int>();var obs=new List<O>();var metrics=new Dictionary<string,double[]>();
    foreach(var x in rows){string dt=(string)x[0],cat=(string)x[4],key=(string)x[3]+"|"+(string)x[5];int di,ci,pi;if(!dateId.TryGetValue(dt,out di)){di=dates.Count;dateId[dt]=di;dates.Add(dt);}if(!catId.TryGetValue(cat,out ci)){ci=cats.Count;catId[cat]=ci;cats.Add(cat);}if(!prodId.TryGetValue(key,out pi)){pi=products.Count;prodId[key]=pi;products.Add(new P((string)x[3],(string)x[5],(string)x[6],(string)x[1],(string)x[7]));}double price=(double)x[2];obs.Add(new O(di,pi,price,ci));string mk=di+"|"+ci;double[] q;if(!metrics.TryGetValue(mk,out q)){q=new double[]{0,0};metrics[mk]=q;}q[0]++;q[1]+=price;}
    using(var w=new StreamWriter(Path.Combine(folder,"dashboard-data.js"),false,new UTF8Encoding(false))){w.Write("window.DT_DATA={\"dates\":[");for(int i=0;i<dates.Count;i++){if(i>0)w.Write(',');w.Write(J(dates[i]));}w.Write("],\"cats\":[");for(int i=0;i<cats.Count;i++){if(i>0)w.Write(',');w.Write(J(cats[i]));}w.Write("],\"products\":[");for(int i=0;i<products.Count;i++){if(i>0)w.Write(',');var p=products[i];w.Write("["+J(p.code)+","+J(p.vmpp)+","+J(p.name)+","+J(p.pack)+","+J(p.form)+"]");}w.Write("],\"obs\":[");for(int i=0;i<obs.Count;i++){if(i>0)w.Write(',');var z=obs[i];w.Write("["+z.d+","+z.p+","+z.v.ToString(CultureInfo.InvariantCulture)+","+z.c+"]");}w.Write("],\"metrics\":[");bool first=true;foreach(var kv in metrics){if(!first)w.Write(',');first=false;var p=kv.Key.Split('|');w.Write("["+p[0]+","+p[1]+","+kv.Value[0]+","+(kv.Value[1]/kv.Value[0]).ToString(CultureInfo.InvariantCulture)+"]");}w.Write("]};");}
    File.WriteAllText(Path.Combine(folder,"build-summary.txt"),"Consolidated records: "+rows.Count+"\r\nHistorical records: "+(h.GetLength(0)-1)+"\r\nJune records added: "+(rows.Count-(h.GetLength(0)-1))+"\r\nDistinct products: "+products.Count+"\r\nCategories: "+String.Join("; ",cats),new UTF8Encoding(true));
  }
}
'@
$xl=New-Object -ComObject Excel.Application;$xl.Visible=$false;$xl.DisplayAlerts=$false
function Get-Grid($file){$wb=$xl.Workbooks.Open((Join-Path $root $file),$null,$true);$x=$wb.Worksheets.Item(1).UsedRange.Value2;$wb.Close($false);return ,$x}
try {$h=Get-Grid 'Drug Tariff data Historical May-26.xlsx';$a=Get-Grid 'Part VIIIA Jun 2026.xlsx';$b=Get-Grid 'Part VIIIB Jun 2026.xlsx';$d=Get-Grid 'Part VIIID Jun 2026.xlsx';$m=Get-Grid 'Cat M prices - Quarter 1 Jun 2026.xlsx';[DTBuilder]::Build($h,$a,$b,$d,$m,$root)} finally {$xl.Quit()}
