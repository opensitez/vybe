// vybe-test: csharp/csharp_switch_expression_core/switch_expr_var_pattern_binds_any_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=42; __Check((o switch{var x when x is int n and n>10=>"big",_=>"other"}).ToString(), "big");
