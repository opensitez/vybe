// vybe-test: csharp/modern_features/typeof_gettype
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

__P((typeof(int).Name).ToString());
__P((typeof(string).Name).ToString());
__P((42.GetType().Name).ToString());
__Check("Int32\nString\nInt32");

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
