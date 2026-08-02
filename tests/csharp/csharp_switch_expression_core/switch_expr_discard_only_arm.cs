// vybe-test: csharp/csharp_switch_expression_core/switch_expr_discard_only_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x=99; __Check((x switch{_=>"always"}).ToString(), "always");
