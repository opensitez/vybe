// vybe-test: csharp/csharp_exception_types/exception_message_survives_catch_and_rethrow_as_inner
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

using static __Harness;

string msg = "";
try {
    try { throw new System.Exception("root"); }
    catch(System.Exception e) { throw new System.Exception("wrap", e); }
}
catch(System.Exception outer) { msg = outer.InnerException.Message; }
__P((msg).ToString());
__Check("root");

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
