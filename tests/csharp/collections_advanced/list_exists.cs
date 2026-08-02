// vybe-test: csharp/collections_advanced/list_exists
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new List<string> { "apple", "banana", "cherry" };
__Check((list.Exists(s => s == "banana")).ToString(), "True");
__Check((list.Exists(s => s == "grape")).ToString(), "False");
