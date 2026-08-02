// vybe-test: csharp/csharp_switch_expressions/switch_expression_uses_default_discard_arm_for_unknown_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 9; __Check((x switch { 1 => "one", 2 => "two", _ => "other" }).ToString(), "other");
