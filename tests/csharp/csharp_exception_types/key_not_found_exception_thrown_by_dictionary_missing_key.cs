// vybe-test: csharp/csharp_exception_types/key_not_found_exception_thrown_by_dictionary_missing_key
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

using static __Harness;

string result = "";
var map = new System.Collections.Generic.Dictionary<string,int>();
try { int v = map["nope"]; }
catch(System.Collections.Generic.KeyNotFoundException) { result = "missing"; }
__P((result).ToString());
__Check("missing");

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
