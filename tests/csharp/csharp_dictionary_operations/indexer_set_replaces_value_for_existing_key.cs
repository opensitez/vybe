// vybe-test: csharp/csharp_dictionary_operations/indexer_set_replaces_value_for_existing_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

using static __Harness;

var d = new System.Collections.Generic.Dictionary<string,int>();
d["x"] = 1;
d["x"] = 9;
__P((d["x"]).ToString());
__Check("9");

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
