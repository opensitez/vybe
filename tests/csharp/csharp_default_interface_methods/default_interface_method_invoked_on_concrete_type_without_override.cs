// vybe-test: csharp/csharp_default_interface_methods/default_interface_method_invoked_on_concrete_type_without_override
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IReporter {
    void Write(string text);
    void Banner(string title) { Write("==" + title + "=="); }
}
class ConsoleReporter : IReporter {
    public void Write(string text) { __Check((text).ToString(), "==start=="); }
}
var reporter = new ConsoleReporter();
reporter.Banner("start");
