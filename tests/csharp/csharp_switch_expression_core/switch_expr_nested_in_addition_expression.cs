// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_in_addition_expression
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=1,b=2; __Check(((a switch{1=>10,_=>0})+(b switch{2=>20,_=>0})).ToString(), "30");
