// vybe-test: csharp/csharp_deferred_execution/where_query_not_executed_until_enumerated
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

using static __Harness;

int count=0;
var q=new[]{1,2,3}
.Where(n=>{count++;return n>1;});
__P((count).ToString());
var list=q.ToList();
__P((count).ToString());
__Check("0\n3");

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
