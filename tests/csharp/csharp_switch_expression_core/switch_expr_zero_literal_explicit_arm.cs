// vybe-test: csharp/csharp_switch_expression_core/switch_expr_zero_literal_explicit_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var n=0; __Check((n switch{0=>"zero",_=>"nz"}).ToString(), "zero");
