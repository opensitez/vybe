// vybe-test: csharp/exceptions_advanced/using_disposes_on_exception
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

try {
    using (var c = new Conn()) {
        throw new Exception("fail");
    }
}
catch (Exception e) {
    __P(("caught: " + e.Message).ToString());
}
__Check("conn closed\ncaught: fail");

class Conn : IDisposable {
    public void Dispose() { __P(("conn closed").ToString()); }
}

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
