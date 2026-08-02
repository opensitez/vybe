// vybe-test: csharp/csharp_collection_initializer_syntax/list_initializer_populates_elements_in_source_order
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var items = new List<int> { 3, 1, 4 };
__Check((items[0]).ToString(), "3");
__Check((items[2]).ToString(), "4");
