// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_guard_string_length_other
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var s="hi"; __Check((s switch{string t when t.Length==4=>"len4",string t=>t.Length.ToString(),_=>"0"}).ToString(), "2");
