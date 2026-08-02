// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_arm_switch_on_inner_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Outer(int n)=>n switch{1=>1 switch{1=>"one-one",_=>"one-other"},2=>"two",_=>"rest"}; __Check((Outer(1)).ToString(), "one-one");
