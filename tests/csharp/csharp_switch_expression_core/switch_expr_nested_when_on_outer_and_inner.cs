// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_when_on_outer_and_inner
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pair=(2,4); __Check((pair switch{(var a,var b) when a<b=>"asc",(var a,var b)=>"desc",_=>"?"}).ToString(), "asc");
