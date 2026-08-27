// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

using static __Harness;

// constructor_overload_surface
var map = new System.Collections.Generic.Dictionary<int, int>();
map[67] = 68;
__P((map.ContainsKey(67) && map[67] == 68).ToString());
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
