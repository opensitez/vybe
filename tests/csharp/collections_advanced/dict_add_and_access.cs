// vybe-test: csharp/collections_advanced/dict_add_and_access
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dict = new Dictionary<string, int>();
dict.Add("one", 1);
dict.Add("two", 2);
dict.Add("three", 3);
__Check((dict["two"]).ToString(), "2");
__Check((dict.Count).ToString(), "3");
