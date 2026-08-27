// vybe-test: csharp/csharp_ref_out_in/out_inline_declaration_in_method_call
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

using static __Harness;

bool ok = int.TryParse("42", out int result);
__P((ok).ToString());
__P((result).ToString());
__Check("True\n42");

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
