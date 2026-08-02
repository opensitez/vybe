// vybe-test: csharp/csharp_switch_expressions/switch_expression_matches_char_literal_cases
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char ch = 'x'; __Check((ch switch { 'x' => "ex", 'y' => "why", _ => "other" }).ToString(), "ex");
