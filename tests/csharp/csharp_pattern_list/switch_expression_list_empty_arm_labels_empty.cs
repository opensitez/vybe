// vybe-test: csharp/csharp_pattern_list/switch_expression_list_empty_arm_labels_empty
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

string Label(int[] a)=>a switch{[]=>"empty",_=>"other"}
;
__P((Label(new int[]{})).ToString());
__Check("empty");

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
