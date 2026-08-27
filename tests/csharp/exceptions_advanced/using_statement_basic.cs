// vybe-test: csharp/exceptions_advanced/using_statement_basic
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

using (var r = new Resource()) {
    __P(("using").ToString());
}
__Check("opened\nusing\ndisposed");

class Resource : IDisposable {
    public Resource() { __P(("opened").ToString()); }
    public void Dispose() { __P(("disposed").ToString()); }
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
