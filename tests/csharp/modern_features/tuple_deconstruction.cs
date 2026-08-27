// vybe-test: csharp/modern_features/tuple_deconstruction
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

var (name, age) = ("Bob", 25);
__P((name).ToString());
__P((age).ToString());
__Check("Bob\n25");

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
