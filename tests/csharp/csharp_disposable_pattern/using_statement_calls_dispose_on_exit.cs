// vybe-test: csharp/csharp_disposable_pattern/using_statement_calls_dispose_on_exit
// origin: languages/csharp/tests/csharp/test_csharp_disposable_pattern.rs

using static __Harness;

var r=new Resource();
using(r){}
__P((r.Disposed).ToString());
__Check("True");

class Resource:System.IDisposable{
    public bool Disposed;
    public void Dispose(){Disposed=true;}
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
