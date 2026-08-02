// vybe-test: csharp/csharp_parsing_formatting/number_format_can_render_percentage_output
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((0.25.ToString("P0")).ToString(), "25 %");
