// vybe-test: csharp/csharp_array_length_variants/array_length_variants_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

using static __Harness;

// array_length_variants
var map = new System.Collections.Generic.Dictionary<int, int>();
map[25] = 26;
__P((map.ContainsKey(25) && map[25] == 26).ToString());
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
