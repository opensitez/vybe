// vybe-test: csharp/csharp_default_interface_methods/default_interface_method_invoked_on_concrete_type_without_override
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods.rs

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

interface IReporter {
    void Write(string text);
    void Banner(string title) { Write("==" + title + "=="); }
}
class ConsoleReporter : IReporter {
    public void Write(string text) { __P((text).ToString()); }
}
var reporter = new ConsoleReporter();
reporter.Banner("start");
__Check("==start==");
