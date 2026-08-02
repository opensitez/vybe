// vybe-test: csharp/csharp_dictionary_contracts/contains_key_reflects_add_and_remove_lifecycle
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var map = new Dictionary<int, string>();
map[1] = "one";
__Check((map.ContainsKey(1) ? "Y" : "N").ToString(), "Y");
map.Remove(1);
__Check((map.ContainsKey(1) ? "Y" : "N").ToString(), "N");
