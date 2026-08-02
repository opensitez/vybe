// vybe-test: csharp/csharp_list_dictionary/list_remove_existing_string_element
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var list = new List<string> { "a", "b" }; list.Remove("a"); __Check((list[0]).ToString(), "b");
