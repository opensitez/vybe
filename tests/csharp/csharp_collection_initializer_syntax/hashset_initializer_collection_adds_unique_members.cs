// vybe-test: csharp/csharp_collection_initializer_syntax/hashset_initializer_collection_adds_unique_members
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var set = new HashSet<int> { 2, 3, 2 };
__Check((set.Count).ToString(), "2");
