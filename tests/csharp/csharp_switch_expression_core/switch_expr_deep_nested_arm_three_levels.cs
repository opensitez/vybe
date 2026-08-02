// vybe-test: csharp/csharp_switch_expression_core/switch_expr_deep_nested_arm_three_levels
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string L(int n)=>n switch{1=>"a",2=>2 switch{2=>"b",3=>"c",_=>"d"},_=>"z"}; __Check((L(2)).ToString(), "b");
