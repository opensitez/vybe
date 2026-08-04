// vybe-test: csharp/csharp_interface_default_impl/default_method_used_when_class_does_not_override
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

interface ILogger{
    void Log(string msg)=>__P(("[LOG] "+msg).ToString());
}
class App:ILogger{}
ILogger app=new App();
app.Log("hello");
__Check("[LOG] hello");
