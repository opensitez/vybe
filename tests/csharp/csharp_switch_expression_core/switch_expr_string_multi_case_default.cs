// vybe-test: csharp/csharp_switch_expression_core/switch_expr_string_multi_case_default
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var key="run"; __Check((key switch{"stop"=>"S","go"=>"G",_=>"?"}).ToString(), "?");
