// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

using static __Harness;

// char_predicate_apis
var map = new System.Collections.Generic.Dictionary<int, int>();
map[23] = 24;
__P((map.ContainsKey(23) && map[23] == 24).ToString());
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
