// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
var values = new System.Collections.Generic.List<int> { 19, 20, 19 }; __Check((values.Count == 3).ToString(), "True");
