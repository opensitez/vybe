// vybe-test: csharp/csharp_switch_expression_core/switch_expr_array_initializer_element
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var codes=new string[]{1 switch{1=>"one",_=>"?"}}; __Check((codes[0]).ToString(), "one");
