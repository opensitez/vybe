// vybe-test: csharp/csharp_pattern_switch_advanced/relational_and_pattern_combines_bounds_check
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

using static __Harness;

string Grade(int n)=>n switch{
    >=90=>"A",
    >=70 and <90=>"B",
    >=50 and <70=>"C",
    _=>"F"}
;
__P((Grade(95)).ToString());
__P((Grade(75)).ToString());
__P((Grade(55)).ToString());
__P((Grade(30)).ToString());
__Check("A\nB\nC\nF");

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
