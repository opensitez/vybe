// vybe-test: csharp/csharp_dictionary_operations/contains_value_finds_value_regardless_of_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

using static __Harness;

var d = new System.Collections.Generic.Dictionary<string,int>{{"a",42}}
;
__P((d.ContainsValue(42)).ToString());
__Check("True");

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
