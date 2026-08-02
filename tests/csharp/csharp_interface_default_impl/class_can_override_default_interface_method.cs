// vybe-test: csharp/csharp_interface_default_impl/class_can_override_default_interface_method
// origin: languages/csharp/tests/csharp/test_csharp_interface_default_impl.rs

interface ILogger{void Log(string msg)=>Console.WriteLine("[LOG] "+msg);}
class SilentApp:ILogger{public void Log(string msg){}}
ILogger app=new SilentApp();
app.Log("hello");
Console.WriteLine("done");
