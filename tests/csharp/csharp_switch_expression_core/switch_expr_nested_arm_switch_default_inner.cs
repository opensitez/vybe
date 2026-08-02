// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_arm_switch_default_inner
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Outer(int n)=>n switch{1=>5 switch{5=>"five",_=>"not-five"},_=>"rest"}; __Check((Outer(1)).ToString(), "not-five");
