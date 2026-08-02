// vybe-test: csharp/csharp_switch_expression_core/switch_expr_interpolated_result_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var score=88; __Check((score switch{>=90=>$"A:{score}",>=80=>$"B:{score}",_=>$"C:{score}"}).ToString(), "B:88");
