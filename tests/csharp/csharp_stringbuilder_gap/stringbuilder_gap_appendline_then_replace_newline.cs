// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_appendline_then_replace_newline
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

using static __Harness;

var sb=new System.Text.StringBuilder();
sb.AppendLine("row");
__P((sb.ToString().Contains("\n")||sb.ToString().Contains("\r")).ToString());
__Check("True");

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
