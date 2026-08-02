// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_and_relational_combo
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var n=18; __Check((n switch{int x when x>=18 and x<21=>"adult",_=>"minor"}).ToString(), "adult");
