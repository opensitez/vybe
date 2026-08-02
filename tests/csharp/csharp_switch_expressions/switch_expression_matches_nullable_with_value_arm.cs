// vybe-test: csharp/csharp_switch_expressions/switch_expression_matches_nullable_with_value_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? value = 12; __Check((value switch { null => "missing", > 10 => "large", _ => "small" }).ToString(), "large");
