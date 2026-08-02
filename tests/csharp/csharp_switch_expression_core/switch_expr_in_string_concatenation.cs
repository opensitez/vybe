// vybe-test: csharp/csharp_switch_expression_core/switch_expr_in_string_concatenation
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var n=4; __Check(("v="+(n switch{4=>"four",_=>"?"})).ToString(), "v=four");
