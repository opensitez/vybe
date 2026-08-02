// vybe-test: csharp/csharp_collections/list_clear
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var list = new List<int>();
list.Add(1);
list.Add(2);
__Check((list.Count).ToString(), "2");
list.Clear();
__Check((list.Count).ToString(), "0");
