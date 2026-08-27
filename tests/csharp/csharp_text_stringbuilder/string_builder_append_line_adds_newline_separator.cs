// vybe-test: csharp/csharp_text_stringbuilder/string_builder_append_line_adds_newline_separator
// origin: languages/csharp/tests/csharp/test_csharp_text_stringbuilder.rs

using static __Harness;

var sb=new System.Text.StringBuilder();
sb.AppendLine("line1").AppendLine("line2");
__P((sb.ToString().Trim().Replace("\r\n","\n")).ToString());
__Check("line1\nline2");

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
