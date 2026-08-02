// vybe-test: csharp/csharp_switch_expressions/switch_expression_handles_tuple_with_discard_component
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pair = (3, 9); __Check((pair switch { (3, _) => "starts-three", (_, 9) => "ends-nine", _ => "other" }).ToString(), "starts-three");
