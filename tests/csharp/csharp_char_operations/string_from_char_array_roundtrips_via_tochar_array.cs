// vybe-test: csharp/csharp_char_operations/string_from_char_array_roundtrips_via_tochar_array
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

using static __Harness;

string s = new string(new char[]{'h','i'});
__P((s).ToString());
__Check("hi");

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
