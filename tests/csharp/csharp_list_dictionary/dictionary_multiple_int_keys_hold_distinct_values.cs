// vybe-test: csharp/csharp_list_dictionary/dictionary_multiple_int_keys_hold_distinct_values
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, int> { [1] = 100, [2] = 200 }; __Check((map[1]).ToString(), "100"); __Check((map[2]).ToString(), "200");
