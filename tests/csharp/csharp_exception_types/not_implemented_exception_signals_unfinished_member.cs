// vybe-test: csharp/csharp_exception_types/not_implemented_exception_signals_unfinished_member
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

using static __Harness;

string result = "";
try { throw new System.NotImplementedException(); }
catch(System.NotImplementedException) { result = "ni"; }
__P((result).ToString());
__Check("ni");

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
