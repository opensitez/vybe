// vybe-test: csharp/csharp_interface_default_impl/class_can_override_default_interface_method
// origin: languages/csharp/tests/csharp/test_csharp_interface_default_impl.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

interface ILogger{void Log(string msg)=>__P(("[LOG] "+msg).ToString());}
class SilentApp:ILogger{public void Log(string msg){}}
ILogger app=new SilentApp();
app.Log("hello");
__P(("done").ToString());
__Check("done");
