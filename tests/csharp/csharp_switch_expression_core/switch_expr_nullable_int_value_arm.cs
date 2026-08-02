// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nullable_int_value_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? v=7; __Check((v switch{null=>"nil",_=>"val"}).ToString(), "val");
