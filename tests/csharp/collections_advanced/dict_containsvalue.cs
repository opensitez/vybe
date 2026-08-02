// vybe-test: csharp/collections_advanced/dict_containsvalue
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dict = new Dictionary<string, int> { { "x", 10 }, { "y", 20 } };
__Check((dict.ContainsValue(10)).ToString(), "True");
__Check((dict.ContainsValue(30)).ToString(), "False");
