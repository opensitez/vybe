// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_indexer_after_mutations
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

using static __Harness;

var sb=new System.Text.StringBuilder("abc");
sb[1]='B';
sb.Append("d");
__P((sb[0]).ToString());
__P((sb.ToString()).ToString());
__Check("a\naBcd");

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
