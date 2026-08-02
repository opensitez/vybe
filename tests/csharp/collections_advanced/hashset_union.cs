// vybe-test: csharp/collections_advanced/hashset_union
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a = new HashSet<int> { 1, 2, 3 };
var b = new HashSet<int> { 3, 4, 5 };
a.UnionWith(b);
__Check((a.Count).ToString(), "5");
