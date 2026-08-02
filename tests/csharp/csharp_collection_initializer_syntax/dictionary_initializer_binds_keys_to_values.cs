// vybe-test: csharp/csharp_collection_initializer_syntax/dictionary_initializer_binds_keys_to_values
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var map = new Dictionary<string, int> { ["x"] = 9, ["y"] = 2 };
__Check((map["y"]).ToString(), "2");
