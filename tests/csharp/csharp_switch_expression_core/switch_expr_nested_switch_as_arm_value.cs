// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_switch_as_arm_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var tier=2; __Check((tier switch{1=>"a",2=>(3 switch{3=>"inner",_=>"outer"}),_=>"?"}).ToString(), "outer");
