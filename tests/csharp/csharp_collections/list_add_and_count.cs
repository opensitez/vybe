// vybe-test: csharp/csharp_collections/list_add_and_count
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var list = new List<int>();
list.Add(10);
list.Add(20);
list.Add(30);
__Check((list.Count).ToString(), "3");
__Check((list[1]).ToString(), "20");
