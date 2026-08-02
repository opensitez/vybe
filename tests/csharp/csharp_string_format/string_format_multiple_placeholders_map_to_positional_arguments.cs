// vybe-test: csharp/csharp_string_format/string_format_multiple_placeholders_map_to_positional_arguments
// origin: languages/csharp/tests/csharp/test_csharp_string_format.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Format("{0} + {1} = {2}", 1, 2, 3)).ToString(), "1 + 2 = 3");
