// vybe-test: csharp/csharp_pattern_list/switch_expression_list_var_pair_arm_returns_sum
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

int SumPair(int[] a)=>a switch{[var x,var y]=>x+y,_=>0}
;
__P((SumPair(new[]{10,20})).ToString());
__Check("30");

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
