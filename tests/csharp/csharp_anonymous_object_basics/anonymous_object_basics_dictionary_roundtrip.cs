// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

using static __Harness;

// anonymous_object_basics
var map = new System.Collections.Generic.Dictionary<int, int>();
map[38] = 39;
__P((map.ContainsKey(38) && map[38] == 39).ToString());
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
