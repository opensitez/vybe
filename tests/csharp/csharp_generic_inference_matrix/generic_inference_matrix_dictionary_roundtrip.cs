// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

using static __Harness;

// generic_inference_matrix
var map = new System.Collections.Generic.Dictionary<int, int>();
map[81] = 82;
__P((map.ContainsKey(81) && map[81] == 82).ToString());
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
