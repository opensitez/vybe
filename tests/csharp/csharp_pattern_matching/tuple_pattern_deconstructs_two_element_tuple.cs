// vybe-test: csharp/csharp_pattern_matching/tuple_pattern_deconstructs_two_element_tuple
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

using static __Harness;

var point = (1, 0);
string axis = point switch {
    (0, 0) => "origin",
    (_, 0) => "x-axis",
    (0, _) => "y-axis",
    _       => "other"
}
;
__P((axis).ToString());
__Check("x-axis");

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
