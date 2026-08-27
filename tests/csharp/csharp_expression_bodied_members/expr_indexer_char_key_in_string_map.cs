// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_char_key_in_string_map
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

var cm = new CharMap();
cm['A'] = 1;
__P((cm['A']).ToString());
__Check("1");

class CharMap { System.Collections.Generic.Dictionary<char, int> m = new(); public int this[char c] { get => m[c]; set => m[c] = value; } }

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
