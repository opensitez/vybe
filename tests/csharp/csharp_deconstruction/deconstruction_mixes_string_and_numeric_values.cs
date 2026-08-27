// vybe-test: csharp/csharp_deconstruction/deconstruction_mixes_string_and_numeric_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

using static __Harness;

var (name, age) = ("Grace", 42);
__P((name + ":" + age).ToString());
__Check("Grace:42");

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
