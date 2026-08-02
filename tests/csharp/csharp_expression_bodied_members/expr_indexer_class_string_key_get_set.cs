// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_string_key_get_set
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Map { System.Collections.Generic.Dictionary<string, int> d = new(); public int this[string k] { get => d[k]; set => d[k] = value; } }
var m = new Map(); m["count"] = 7; __Check((m["count"]).ToString(), "7");
