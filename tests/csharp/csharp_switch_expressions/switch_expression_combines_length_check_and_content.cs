// vybe-test: csharp/csharp_switch_expressions/switch_expression_combines_length_check_and_content
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var text = "tool"; __Check((text switch { string s when s.Length == 4 => "len4", string s => s, _ => "none" }).ToString(), "len4");
