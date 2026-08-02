// vybe-test: csharp/collections_advanced/dict_remove
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dict = new Dictionary<string, int> { { "a", 1 }, { "b", 2 }, { "c", 3 } };
dict.Remove("b");
__Check((dict.Count).ToString(), "2");
__Check((dict.ContainsKey("b")).ToString(), "False");
