// vybe-test: csharp/csharp_list_dictionary/dictionary_int_keys_store_negative_numbers
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, int> { [-1] = 100 }; __Check((map[-1]).ToString(), "100");
