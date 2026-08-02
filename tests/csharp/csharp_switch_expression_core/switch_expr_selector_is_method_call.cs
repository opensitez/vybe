// vybe-test: csharp/csharp_switch_expression_core/switch_expr_selector_is_method_call
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Twice(int x)=>x*2; __Check((Twice(3) switch{6=>"ok",_=>"no"}).ToString(), "ok");
