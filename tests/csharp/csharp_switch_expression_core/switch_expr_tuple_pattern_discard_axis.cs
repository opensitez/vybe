// vybe-test: csharp/csharp_switch_expression_core/switch_expr_tuple_pattern_discard_axis
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var p=(3,0); __Check((p switch{(0,0)=>"origin",(_,0)=>"x",_=>"away"}).ToString(), "x");
