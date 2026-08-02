// vybe-test: csharp/csharp_switch_expression_core/switch_expr_type_pattern_int_increment
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=6; __Check((o switch{int n=>(n+1).ToString(),_=>"?"}).ToString(), "7");
