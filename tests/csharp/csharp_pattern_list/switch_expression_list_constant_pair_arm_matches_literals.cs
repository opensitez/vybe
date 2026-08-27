// vybe-test: csharp/csharp_pattern_list/switch_expression_list_constant_pair_arm_matches_literals
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

string Code(int[] a)=>a switch{[1,2]=>"twelve",_=>"other"}
;
__P((Code(new[]{1,2})).ToString());
__Check("twelve");

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
