// vybe-test: csharp/csharp_switch_expressions/switch_expression_uses_when_guard_for_odd_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 5; __Check((x switch { int n when n % 2 == 0 => "even", int n => "odd" }).ToString(), "odd");
