// vybe-test: csharp/csharp_switch_expression_core/switch_expr_tuple_pattern_origin
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var p=(0,0); __Check((p switch{(0,0)=>"origin",_=>"away"}).ToString(), "origin");
