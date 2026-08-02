// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_char_key_in_string_map
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class CharMap { System.Collections.Generic.Dictionary<char, int> m = new(); public int this[char c] { get => m[c]; set => m[c] = value; } }
var cm = new CharMap(); cm['A'] = 1; __Check((cm['A']).ToString(), "1");
