// vybe-test: csharp/csharp_switch_expression_core/switch_expr_arm_calls_method
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Double(int x)=>x*2; __Check((5 switch{5=>Double(5),_=>0}).ToString(), "10");
