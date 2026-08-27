// vybe-test: csharp/csharp_pattern_list/switch_expression_list_when_guard_on_vars
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

string Rank(int[] a)=>a switch{[var x,var y] when x>y=>"desc",[var x,var y]=>"asc",_=>"other"}
;
__P((Rank(new[]{5,2})).ToString());
__Check("desc");

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
