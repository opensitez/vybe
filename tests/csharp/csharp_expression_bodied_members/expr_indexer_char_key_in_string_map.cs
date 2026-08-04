// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_char_key_in_string_map
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class CharMap { System.Collections.Generic.Dictionary<char, int> m = new(); public int this[char c] { get => m[c]; set => m[c] = value; } }
var cm = new CharMap(); cm['A'] = 1; __P((cm['A']).ToString());
__Check("1");
