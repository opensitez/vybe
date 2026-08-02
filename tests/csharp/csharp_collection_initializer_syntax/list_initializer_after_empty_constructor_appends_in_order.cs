// vybe-test: csharp/csharp_collection_initializer_syntax/list_initializer_after_empty_constructor_appends_in_order
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var items = new List<string>();
items.Add("first");
items.Add("second");
__Check((items[1]).ToString(), "second");
