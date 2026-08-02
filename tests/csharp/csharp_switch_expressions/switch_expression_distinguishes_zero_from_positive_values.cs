// vybe-test: csharp/csharp_switch_expressions/switch_expression_distinguishes_zero_from_positive_values
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 0; __Check((x switch { < 0 => "neg", 0 => "zero", > 0 => "pos" }).ToString(), "zero");
