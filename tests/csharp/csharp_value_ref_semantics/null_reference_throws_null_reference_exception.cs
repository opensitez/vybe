// vybe-test: csharp/csharp_value_ref_semantics/null_reference_throws_null_reference_exception
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

using static __Harness;

string r="";
try{string s=null;int len=s.Length;}
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
