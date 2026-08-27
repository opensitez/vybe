// vybe-test: csharp/csharp_disposable_pattern/using_declaration_disposes_at_end_of_block
// origin: languages/csharp/tests/csharp/test_csharp_disposable_pattern.rs

using static __Harness;

R r;
{using var x=new R(); r=x;}
__P((r.Gone).ToString());
__Check("True");

class R:System.IDisposable{public bool Gone;public void Dispose(){Gone=true;}}

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
