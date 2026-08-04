// vybe-test: csharp/exceptions_advanced/using_statement_basic
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

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

class Resource : IDisposable {
    public Resource() { __P(("opened").ToString()); }
    public void Dispose() { __P(("disposed").ToString()); }
}
using (var r = new Resource()) {
    __P(("using").ToString());
}
__Check("opened\nusing\ndisposed");
