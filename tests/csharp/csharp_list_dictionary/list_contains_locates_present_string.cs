// vybe-test: csharp/csharp_list_dictionary/list_contains_locates_present_string
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<string> { "cat" }; __Check((list.Contains("cat")).ToString(), "True");
