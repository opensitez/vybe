// vybe-test: csharp/type_features/multi_var_no_initializer
// origin: languages/csharp/tests/csharp/test_type_features.rs

using static __Harness;

string a = "hello", b = "world";
__P((a + " " + b).ToString());
__Check("hello world");

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
