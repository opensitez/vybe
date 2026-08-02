// vybe-test: csharp/csharp_collections/list_contains_remove
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var list = new List<string>();
list.Add("apple");
list.Add("banana");
list.Add("cherry");
__Check((list.Contains("banana")).ToString(), "True");
list.Remove("banana");
__Check((list.Count).ToString(), "2");
__Check((list.Contains("banana")).ToString(), "False");
