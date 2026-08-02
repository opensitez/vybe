// vybe-test: csharp/csharp_switch_expression_core/switch_expr_type_pattern_string_upper
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o="abc"; __Check((o switch{string s=>s.ToUpper(),_=>"?"}).ToString(), "ABC");
