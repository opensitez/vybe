// vybe-test: csharp/csharp_disposable_pattern/try_finally_equivalent_to_using_for_cleanup
// origin: languages/csharp/tests/csharp/test_csharp_disposable_pattern.rs

using static __Harness;

bool cleaned=false;
var f=new System.Action(()=>cleaned=true);
try{}
finally{f();}
__P((cleaned).ToString());
__Check("True");

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
