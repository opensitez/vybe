// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_guard_false_skips_to_next_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x=3; __Check((x switch{int n when n>10=>"big",int n when n>1=>"mid",_=>"small"}).ToString(), "mid");
