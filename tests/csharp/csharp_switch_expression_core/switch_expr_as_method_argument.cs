// vybe-test: csharp/csharp_switch_expression_core/switch_expr_as_method_argument
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Show(string s){__Check((s).ToString(), "three");} Show(3 switch{3=>"three",_=>"other"});
