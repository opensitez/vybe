// vybe-test: csharp/csharp_linq_let_join/multiple_from_clauses_produce_cartesian_product
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

using static __Harness;

var result=from a in new[]{1,2}
from b in new[]{10,20}
select a*b;
int sum=0;
foreach(var x in result) sum+=x;
__P((sum).ToString());
__Check("90");

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
