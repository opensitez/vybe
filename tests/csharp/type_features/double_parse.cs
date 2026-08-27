// vybe-test: csharp/type_features/double_parse
// origin: languages/csharp/tests/csharp/test_type_features.rs

using static __Harness;

__P((double.Parse("3.14")).ToString());
__Check("3.14");

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
