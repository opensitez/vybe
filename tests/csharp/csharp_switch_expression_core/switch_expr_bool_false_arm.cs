// vybe-test: csharp/csharp_switch_expression_core/switch_expr_bool_false_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool flag=false; __Check((flag switch{true=>"yes",false=>"no"}).ToString(), "no");
