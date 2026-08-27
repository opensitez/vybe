// vybe-test: csharp/csharp_linq_let_join/into_continues_query_after_group
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

using static __Harness;

var nums=new[]{1,2,3,4,5,6}
;
var q=from n in nums
      group n by n%2 into g
      select g.Key;
var keys=q.OrderBy(x=>x).ToList();
__P((keys[0]).ToString());
__P((keys[1]).ToString());
__Check("0\n1");

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
