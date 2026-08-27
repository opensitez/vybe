// vybe-test: csharp/csharp_array_apis/array_empty_returns_zero_length_array
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

using static __Harness;

__P((System.Array.Empty<string>().Length).ToString());
__Check("0");

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
