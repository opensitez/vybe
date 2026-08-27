// vybe-test: csharp/csharp_interface_contracts/idisposable_dispose_called_by_using_statement
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

using static __Harness;

using(var r = new Resource()) { }
__P((Resource.Disposed).ToString());
__Check("1");

class Resource : System.IDisposable {
    public static int Disposed = 0;
    public void Dispose() => Disposed++;
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
