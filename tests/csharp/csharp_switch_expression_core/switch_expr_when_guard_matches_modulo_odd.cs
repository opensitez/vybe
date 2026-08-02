// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_guard_matches_modulo_odd
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x=11; __Check((x switch{int n when n%2==0=>"even",int n=>"odd"}).ToString(), "odd");
