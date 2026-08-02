// vybe-test: csharp/collections_advanced/collection_initializer_list
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var names = new List<string> { "Alice", "Bob", "Charlie" };
__Check((names.Count).ToString(), "3");
__Check((names[1]).ToString(), "Bob");
