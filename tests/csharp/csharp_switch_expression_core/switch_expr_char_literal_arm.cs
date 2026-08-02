// vybe-test: csharp/csharp_switch_expression_core/switch_expr_char_literal_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char c='q'; __Check((c switch{'q'=>"que",_=>"other"}).ToString(), "que");
