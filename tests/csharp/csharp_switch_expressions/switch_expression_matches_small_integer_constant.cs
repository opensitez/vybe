// vybe-test: csharp/csharp_switch_expressions/switch_expression_matches_small_integer_constant
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 2; __Check((x switch { 1 => "one", 2 => "two", _ => "other" }).ToString(), "two");
