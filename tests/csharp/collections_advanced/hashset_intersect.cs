// vybe-test: csharp/collections_advanced/hashset_intersect
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a = new HashSet<int> { 1, 2, 3, 4 };
var b = new HashSet<int> { 2, 4, 6 };
a.IntersectWith(b);
__Check((a.Count).ToString(), "2");
__Check((a.Contains(2)).ToString(), "True");
__Check((a.Contains(4)).ToString(), "True");
