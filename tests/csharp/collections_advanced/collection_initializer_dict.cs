// vybe-test: csharp/collections_advanced/collection_initializer_dict
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ages = new Dictionary<string, int> {
    { "Alice", 30 },
    { "Bob", 25 }
};
__Check((ages["Alice"]).ToString(), "30");
__Check((ages.Count).ToString(), "2");
