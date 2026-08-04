// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_string_key_get_set
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

class Map { System.Collections.Generic.Dictionary<string, int> d = new(); public int this[string k] { get => d[k]; set => d[k] = value; } }
var m = new Map(); m["count"] = 7; __P((m["count"]).ToString());
__Check("7");
