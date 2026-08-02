// vybe-test: csharp/csharp_switch_expression_core/switch_expr_negative_int_literal_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var n=-2; __Check((n switch{-2=>"neg-two",_=>"other"}).ToString(), "neg-two");
