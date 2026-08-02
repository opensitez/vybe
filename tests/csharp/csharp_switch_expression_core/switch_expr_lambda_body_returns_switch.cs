// vybe-test: csharp/csharp_switch_expression_core/switch_expr_lambda_body_returns_switch
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var fn=(int x)=>x switch{0=>"z",_=>"nz"}; __Check((fn(0)).ToString(), "z");
