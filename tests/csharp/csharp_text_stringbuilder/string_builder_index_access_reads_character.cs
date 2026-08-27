// vybe-test: csharp/csharp_text_stringbuilder/string_builder_index_access_reads_character
// origin: languages/csharp/tests/csharp/test_csharp_text_stringbuilder.rs

using static __Harness;

var sb=new System.Text.StringBuilder("hello");
__P((sb[1]).ToString());
__Check("e");

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
