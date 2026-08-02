// vybe-test: csharp/csharp_parsing_formatting/convert_to_string_can_render_integer_value
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Convert.ToString(25)).ToString(), "25");
