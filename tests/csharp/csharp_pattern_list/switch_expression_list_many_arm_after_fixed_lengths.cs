// vybe-test: csharp/csharp_pattern_list/switch_expression_list_many_arm_after_fixed_lengths
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

string Size(int[] a)=>a switch{[]=>"0",[_]=>"1",[_,_]=>"2",_=>"many"}
;
__P((Size(new[]{1,2,3})).ToString());
__Check("many");

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
