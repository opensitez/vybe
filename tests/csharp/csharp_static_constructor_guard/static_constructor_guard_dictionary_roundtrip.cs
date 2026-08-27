// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

using static __Harness;

// static_constructor_guard
var map = new System.Collections.Generic.Dictionary<int, int>();
map[69] = 70;
__P((map.ContainsKey(69) && map[69] == 70).ToString());
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
