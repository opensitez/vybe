// vybe-test: csharp/csharp_switch_expression_core/switch_expr_or_constant_arms_weekday
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var day="Monday"; __Check((day switch{"Saturday" or "Sunday"=>"off",_=>"work"}).ToString(), "work");
