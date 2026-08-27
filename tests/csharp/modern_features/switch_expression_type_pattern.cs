// vybe-test: csharp/modern_features/switch_expression_type_pattern
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

object obj = 42;
string result = obj switch {
    int i => "int: " + i,
    string s => "string: " + s,
    _ => "unknown"
}
;
__P((result).ToString());
__Check("int: 42");

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
