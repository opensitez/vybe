// vybe-test: csharp/csharp_switch_expression_core/switch_expr_assign_to_local_int_literal_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var code=2; var label=code switch{1=>"one",2=>"two",_=>"many"}; __Check((label).ToString(), "two");
