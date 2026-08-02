// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_guard_first_true_stops
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x=15; __Check((x switch{int n when n>10=>"big",int n when n>1=>"mid",_=>"small"}).ToString(), "big");
