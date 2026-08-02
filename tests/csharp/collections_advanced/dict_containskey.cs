// vybe-test: csharp/collections_advanced/dict_containskey
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dict = new Dictionary<string, int> { { "a", 1 }, { "b", 2 } };
__Check((dict.ContainsKey("a")).ToString(), "True");
__Check((dict.ContainsKey("c")).ToString(), "False");
