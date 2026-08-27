// vybe-test: csharp/csharp_nullable_ref_types/null_forgiving_operator_suppresses_warning_still_nulls_at_runtime
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

using static __Harness;

string? s=null;
string r="ok";
try{__P((s!.Length).ToString());}
catch(System.NullReferenceException){r="null";}
__P((r).ToString());
__Check("null");

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
