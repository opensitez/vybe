// vybe-test: csharp/csharp_switch_expressions/switch_expression_orders_tuple_pattern_arms
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pair = (1, 0); __Check((pair switch { (0, 0) => "origin", (1, 0) => "unit-x", _ => "other" }).ToString(), "unit-x");
