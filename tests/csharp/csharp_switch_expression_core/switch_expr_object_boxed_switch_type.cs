// vybe-test: csharp/csharp_switch_expression_core/switch_expr_object_boxed_switch_type
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=12L; __Check((o switch{long l=>l.ToString(),int i=>i.ToString(),_=>"?"}).ToString(), "12");
