// vybe-test: csharp/csharp_switch_expression_core/switch_expr_double_nested_in_arm_result
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Pick(int a,int b)=>a switch{1=>b switch{2=>10,3=>20,_=>0},_=>-1}; __Check((Pick(1,3)).ToString(), "20");
