// vybe-test: csharp/csharp_collections/list_sort_and_reverse
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var list = new List<int>();
list.Add(3);
list.Add(1);
list.Add(4);
list.Add(1);
list.Add(5);
list.Sort();
__Check((list[0]).ToString(), "1");
__Check((list[4]).ToString(), "5");
list.Reverse();
__Check((list[0]).ToString(), "5");
