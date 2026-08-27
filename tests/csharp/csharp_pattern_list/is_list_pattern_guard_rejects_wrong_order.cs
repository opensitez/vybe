// vybe-test: csharp/csharp_pattern_list/is_list_pattern_guard_rejects_wrong_order
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

int[] data=new[]{8,4}
;
if(data is [var a,var b] && a<b) __P(("ordered").ToString());
else __P(("not").ToString());
__Check("not");

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
