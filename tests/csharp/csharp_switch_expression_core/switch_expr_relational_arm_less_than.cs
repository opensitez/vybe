// vybe-test: csharp/csharp_switch_expression_core/switch_expr_relational_arm_less_than
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var n=2; __Check((n switch{<5=>"low",_=>"high"}).ToString(), "low");
