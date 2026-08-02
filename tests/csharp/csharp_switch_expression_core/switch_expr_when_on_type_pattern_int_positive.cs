// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_on_type_pattern_int_positive
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=8; __Check((o switch{int n when n>0=>"pos",_=>"other"}).ToString(), "pos");
