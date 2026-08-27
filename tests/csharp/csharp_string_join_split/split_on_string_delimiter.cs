// vybe-test: csharp/csharp_string_join_split/split_on_string_delimiter
// origin: languages/csharp/tests/csharp/test_csharp_string_join_split.rs

using static __Harness;

var parts="one::two::three".Split("::");
__P((parts.Length).ToString());
__P((parts[2]).ToString());
__Check("3\nthree");

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
