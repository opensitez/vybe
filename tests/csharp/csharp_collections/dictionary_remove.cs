// vybe-test: csharp/csharp_collections/dictionary_remove
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var dict = new Dictionary<string, int>();
dict.Add("a", 1);
dict.Add("b", 2);
dict.Remove("a");
__Check((dict.Count).ToString(), "1");
