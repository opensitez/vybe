// vybe-test: csharp/csharp_switch_expressions/switch_expression_matches_object_type_pattern
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object item = "hello"; __Check((item switch { string text => text.ToUpper(), int number => (number * 2).ToString(), _ => "other" }).ToString(), "HELLO");
