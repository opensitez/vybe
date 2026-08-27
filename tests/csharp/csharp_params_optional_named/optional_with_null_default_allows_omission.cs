// vybe-test: csharp/csharp_params_optional_named/optional_with_null_default_allows_omission
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

using static __Harness;

string Label(string text, string tag=null) => tag==null?text:$"[{tag}]{text}";
__P((Label("msg")).ToString());
__P((Label("msg","info")).ToString());
__Check("msg\n[info]msg");

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
