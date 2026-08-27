// vybe-test: csharp/csharp_abstract_class_matrix/abstract_class_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class_matrix.rs

using static __Harness;

// abstract_class_matrix
var map = new System.Collections.Generic.Dictionary<int, int>();
map[72] = 73;
__P((map.ContainsKey(72) && map[72] == 73).ToString());
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
