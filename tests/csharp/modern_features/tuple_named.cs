// vybe-test: csharp/modern_features/tuple_named
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

var p = (Name: "Alice", Age: 30);
__P((p.Name).ToString());
__P((p.Age).ToString());
__Check("Alice\n30");

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
