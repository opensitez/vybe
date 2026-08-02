// vybe-test: csharp/csharp_collections/dictionary_basic
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var dict = new Dictionary<string, int>();
dict.Add("x", 10);
dict.Add("y", 20);
__Check((dict["x"]).ToString(), "10");
__Check((dict.ContainsKey("y")).ToString(), "True");
__Check((dict.ContainsKey("z")).ToString(), "False");
__Check((dict.Count).ToString(), "2");
