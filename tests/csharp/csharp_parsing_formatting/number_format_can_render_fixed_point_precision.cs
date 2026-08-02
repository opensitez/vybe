// vybe-test: csharp/csharp_parsing_formatting/number_format_can_render_fixed_point_precision
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((3.14159.ToString("F2")).ToString(), "3.14");
