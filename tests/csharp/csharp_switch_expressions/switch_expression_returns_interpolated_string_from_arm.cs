// vybe-test: csharp/csharp_switch_expressions/switch_expression_returns_interpolated_string_from_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var score = 87; __Check((score switch { >= 90 => $"A:{score}", >= 80 => $"B:{score}", _ => $"C:{score}" }).ToString(), "B:87");
