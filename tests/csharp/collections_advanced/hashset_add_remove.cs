// vybe-test: csharp/collections_advanced/hashset_add_remove
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var set = new HashSet<string>();
set.Add("apple");
set.Add("banana");
set.Add("apple");
__Check((set.Count).ToString(), "2");
set.Remove("apple");
__Check((set.Count).ToString(), "1");
__Check((set.Contains("apple")).ToString(), "False");
