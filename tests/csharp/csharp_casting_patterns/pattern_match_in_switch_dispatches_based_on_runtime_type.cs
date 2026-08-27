// vybe-test: csharp/csharp_casting_patterns/pattern_match_in_switch_dispatches_based_on_runtime_type
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

using static __Harness;

object o=42;
string r=o switch{int n=>$"int:{n}",string s=>$"str:{s}",_=>"other"}
;
__P((r).ToString());
__Check("int:42");

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
