// vybe-test: csharp/csharp_boolean_algebra_surfaces/boolean_algebra_surfaces_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_boolean_algebra_surfaces.rs

using static __Harness;

// boolean_algebra_surfaces
var map = new System.Collections.Generic.Dictionary<int, int>();
map[11] = 12;
__P((map.ContainsKey(11) && map[11] == 12).ToString());
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
