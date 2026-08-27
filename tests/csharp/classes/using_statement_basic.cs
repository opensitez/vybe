// vybe-test: csharp/classes/using_statement_basic
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

using (var r = new Resource()) {
    __P("InUsing");
}
__Check("InUsing\nDisposed");

class Resource : System.IDisposable {
    public void Dispose() => __P("Disposed");
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
