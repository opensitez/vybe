// vybe-test: csharp/csharp_string_advanced_ops/string_replace_specific_occurrence_via_stringbuilder
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

using static __Harness;

string s="aababc";
var sb=new System.Text.StringBuilder(s);
int idx=s.IndexOf("ab",1);
sb.Remove(idx,2).Insert(idx,"XX");
__P((sb.ToString()).ToString());
__Check("aXXabc");

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
