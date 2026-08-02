// vybe-test: csharp/csharp_string_methods/concat_static_joins_array_of_strings_without_separator
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Concat("a","b","c")).ToString(), "abc");
