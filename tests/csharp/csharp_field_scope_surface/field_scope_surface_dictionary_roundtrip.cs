// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

using static __Harness;

// field_scope_surface
var map = new System.Collections.Generic.Dictionary<int, int>();
map[63] = 64;
__P((map.ContainsKey(63) && map[63] == 64).ToString());
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
