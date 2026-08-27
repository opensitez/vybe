// vybe-test: csharp/csharp_interface_default_impl/class_can_override_default_interface_method
// origin: languages/csharp/tests/csharp/test_csharp_interface_default_impl.rs

using static __Harness;

ILogger app=new SilentApp();
app.Log("hello");
__P(("done").ToString());
__Check("done");

interface ILogger{void Log(string msg)=>__P(("[LOG] "+msg).ToString());}

class SilentApp:ILogger{public void Log(string msg){}}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
