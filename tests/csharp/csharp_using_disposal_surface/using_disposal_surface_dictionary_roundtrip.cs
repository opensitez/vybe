// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

using static __Harness;

// using_disposal_surface
var map = new System.Collections.Generic.Dictionary<int, int>();
map[52] = 53;
__P((map.ContainsKey(52) && map[52] == 53).ToString());
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
