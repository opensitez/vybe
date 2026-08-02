// vybe-test: csharp/csharp_switch_expressions/switch_expression_handles_negative_number_with_relational_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = -3; __Check((x switch { < 0 => "neg", 0 => "zero", > 0 => "pos" }).ToString(), "neg");
