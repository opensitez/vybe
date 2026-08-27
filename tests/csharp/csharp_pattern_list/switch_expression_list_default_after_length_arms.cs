// vybe-test: csharp/csharp_pattern_list/switch_expression_list_default_after_length_arms
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

string Bucket(int[] a)=>a switch{[]=>"e",[_]=>"s",_=>"m"}
;
__P((Bucket(new[]{1,2})).ToString());
__Check("m");

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
