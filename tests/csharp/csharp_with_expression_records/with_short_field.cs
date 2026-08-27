// vybe-test: csharp/csharp_with_expression_records/with_short_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var t=(new ShortBox(1)) with{S=1000}
;
__P((t.S).ToString());
__Check("1000");

record ShortBox(short S);

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
