// vybe-test: csharp/csharp_bcl_collections/sorted_dictionary_enumerator_yields_keys_in_ascending_order
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

using static __Harness;

var map = new System.Collections.Generic.SortedDictionary<int, string>();
map[3] = "c";
map[1] = "a";
int firstKey = 0;
foreach (var pair in map) { firstKey = pair.Key; break; }
__P((firstKey).ToString());
__Check("1");

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
