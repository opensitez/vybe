// vybe-test: csharp/csharp_switch_expression_core/switch_expr_multiple_commas_trailing_ok
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var n=1; __Check((n switch{1=>"one",2=>"two",_=>"many",}).ToString(), "one");
