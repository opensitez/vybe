// vybe-test: csharp/collections/list_index_access
// origin: languages/csharp/tests/csharp/test_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new List<int>();
        list.Add(10);
        list.Add(20);
        list.Add(30);
        __Check((list[0]).ToString(), "10");
        __Check((list[2]).ToString(), "30");
