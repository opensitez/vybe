// vybe-test: csharp/csharp_switch_expressions/switch_expression_matches_boolean_tuple_combinations
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flags = (true, false); __Check((flags switch { (true, true) => "both", (true, false) => "left", _ => "other" }).ToString(), "left");
