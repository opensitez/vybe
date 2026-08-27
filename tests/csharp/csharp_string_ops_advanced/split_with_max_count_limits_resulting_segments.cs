// vybe-test: csharp/csharp_string_ops_advanced/split_with_max_count_limits_resulting_segments
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

using static __Harness;

var parts="a:b:c:d".Split(':',2);
__P((parts.Length).ToString());
__P((parts[1]).ToString());
__Check("2\nb:c:d");

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
