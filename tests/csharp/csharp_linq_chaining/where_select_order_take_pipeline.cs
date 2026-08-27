// vybe-test: csharp/csharp_linq_chaining/where_select_order_take_pipeline
// origin: languages/csharp/tests/csharp/test_csharp_linq_chaining.rs

using static __Harness;

var result=new[]{5,3,8,1,9,2,7,4,6}
.Where(n=>n>3)
    .Select(n=>n*n)
    .OrderBy(n=>n)
    .Take(3);
foreach(var n in result) __P((n).ToString());
__Check("16\n25\n36");

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
