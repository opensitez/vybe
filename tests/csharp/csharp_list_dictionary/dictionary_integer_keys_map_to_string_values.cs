// vybe-test: csharp/csharp_list_dictionary/dictionary_integer_keys_map_to_string_values
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, string> { [10] = "ten", [20] = "twenty" }; __Check((map[20]).ToString(), "twenty");
