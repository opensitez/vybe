// vybe-test: csharp/collections_advanced/hashset_basic
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var set = new HashSet<int> { 1, 2, 3, 2, 1 };
__Check((set.Count).ToString(), "3");
__Check((set.Contains(2)).ToString(), "True");
__Check((set.Contains(5)).ToString(), "False");
