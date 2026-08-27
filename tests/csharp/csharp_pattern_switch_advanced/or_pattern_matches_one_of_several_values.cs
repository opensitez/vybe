// vybe-test: csharp/csharp_pattern_switch_advanced/or_pattern_matches_one_of_several_values
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

using static __Harness;

string Weekend(string day)=>day switch{
    "Saturday" or "Sunday"=>"weekend",
    _=>"weekday"}
;
__P((Weekend("Saturday")).ToString());
__P((Weekend("Monday")).ToString());
__Check("weekend\nweekday");

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
