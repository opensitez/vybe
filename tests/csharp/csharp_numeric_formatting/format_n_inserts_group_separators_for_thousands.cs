// vybe-test: csharp/csharp_numeric_formatting/format_n_inserts_group_separators_for_thousands
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

using static __Harness;

var s = (1234567).ToString("N0",
    System.Globalization.CultureInfo.InvariantCulture);
__P((s).ToString());
__Check("1,234,567");

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
