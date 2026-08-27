// vybe-test: csharp/csharp_linq_skip_take_distinct/paging_skip_take_repeat_page_two_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

using static __Harness;

var src=new[]{1,2,3,4,5,6,7,8,9}
;
var page=src.Skip(3).Take(3);
__P((page.Count()).ToString());
__Check("3");

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
