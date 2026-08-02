// vybe-test: csharp/csharp_collections/list_indexof
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var list = new List<string>();
list.Add("a");
list.Add("b");
list.Add("c");
__Check((list.IndexOf("b")).ToString(), "1");
__Check((list.IndexOf("z")).ToString(), "-1");
