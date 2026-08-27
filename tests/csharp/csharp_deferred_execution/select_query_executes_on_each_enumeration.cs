// vybe-test: csharp/csharp_deferred_execution/select_query_executes_on_each_enumeration
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

using static __Harness;

int calls=0;
var q=new[]{1,2,3}
.Select(n=>{calls++;return n*2;});
var r1=q.ToList();
var r2=q.ToList();
__P((calls).ToString());
__Check("6");

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
