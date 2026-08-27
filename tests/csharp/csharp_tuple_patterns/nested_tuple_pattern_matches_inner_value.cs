// vybe-test: csharp/csharp_tuple_patterns/nested_tuple_pattern_matches_inner_value
// origin: languages/csharp/tests/csharp/test_csharp_tuple_patterns.rs

using static __Harness;

var data=((1,2),(3,4));
var((a,b),(c,d))=data;
__P((a+b+c+d).ToString());
__Check("10");

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
