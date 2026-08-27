// vybe-test: csharp/csharp_pattern_list/is_list_string_var_pattern_captures_element
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

string[] words=new[]{"hi"}
;
if(words is [var w]) __P((w).ToString());
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
