// vybe-test: csharp/collections_advanced/list_trueforall
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new List<int> { 2, 4, 6, 8 };
__Check((list.TrueForAll(x => x % 2 == 0)).ToString(), "True");
list.Add(3);
__Check((list.TrueForAll(x => x % 2 == 0)).ToString(), "False");
