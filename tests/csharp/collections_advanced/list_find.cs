// vybe-test: csharp/collections_advanced/list_find
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new List<int> { 1, 2, 3, 4, 5 };
var found = list.Find(x => x > 3);
__Check((found).ToString(), "4");
