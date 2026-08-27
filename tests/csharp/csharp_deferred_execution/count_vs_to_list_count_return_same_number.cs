// vybe-test: csharp/csharp_deferred_execution/count_vs_to_list_count_return_same_number
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

using static __Harness;

var q=new[]{1,2,3,4}
.Where(x=>x%2==0);
__P((q.Count()).ToString());
__P((q.ToList().Count).ToString());
__Check("2\n2");

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
