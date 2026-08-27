// vybe-test: csharp/csharp_datetime_formatting/tostring_with_dd_slash_mm_slash_yyyy_format
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

using static __Harness;

var d = new System.DateTime(2024,1,5);
__P((d.ToString("dd/MM/yyyy")).ToString());
__Check("05/01/2024");

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
