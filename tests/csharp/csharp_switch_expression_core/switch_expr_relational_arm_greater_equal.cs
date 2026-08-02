// vybe-test: csharp/csharp_switch_expression_core/switch_expr_relational_arm_greater_equal
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var n=10; __Check((n switch{>=10=>"ten+",_=>"low"}).ToString(), "ten+");
