// vybe-test: csharp/exceptions_advanced/using_disposes_on_exception
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

class Conn : IDisposable {
    public void Dispose() { __P(("conn closed").ToString()); }
}
try {
    using (var c = new Conn()) {
        throw new Exception("fail");
    }
} catch (Exception e) {
    __P(("caught: " + e.Message).ToString());
}
__Check("conn closed\ncaught: fail");
