// vybe-test: csharp/csharp_switch_expressions/switch_expression_uses_parenthesized_result_expression
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 4; __Check((x switch { 4 => (2 + 3), _ => 0 }).ToString(), "5");
