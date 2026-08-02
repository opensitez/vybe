// vybe-test: csharp/csharp_class_indexers/string_keyed_indexer_stores_and_retrieves_values
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bag {
    System.Collections.Generic.Dictionary<string, int> map = new();
    public int this[string key] {
        get { return map[key]; }
        set { map[key] = value; }
    }
}
var bag = new Bag();
bag["count"] = 7;
__Check((bag["count"]).ToString(), "7");
