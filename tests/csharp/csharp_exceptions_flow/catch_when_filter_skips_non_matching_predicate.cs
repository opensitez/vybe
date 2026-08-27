// vybe-test: csharp/csharp_exceptions_flow/catch_when_filter_skips_non_matching_predicate
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_flow.rs

using static __Harness;

string r="unhandled";
try{throw new System.Exception("skip");}
catch(System.Exception ex) when(ex.Message=="match"){r="matched";}
catch(System.Exception){r="caught";}
__P((r).ToString());
__Check("caught");

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
