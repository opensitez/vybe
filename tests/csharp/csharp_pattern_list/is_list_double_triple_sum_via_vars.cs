// vybe-test: csharp/csharp_pattern_list/is_list_double_triple_sum_via_vars
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

double[] vals=new[]{1.5,2.0,2.5}
;
if(vals is [var a,var b,var c]) __P((a+b+c).ToString());
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
