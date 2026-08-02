// vybe-test: csharp/csharp_interface_default_impl/default_method_used_when_class_does_not_override
// origin: languages/csharp/tests/csharp/test_csharp_interface_default_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ILogger{
    void Log(string msg)=>__Check(("[LOG] "+msg).ToString(), "[LOG] hello");
}
class App:ILogger{}
ILogger app=new App();
app.Log("hello");
