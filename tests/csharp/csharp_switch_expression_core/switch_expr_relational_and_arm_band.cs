// vybe-test: csharp/csharp_switch_expression_core/switch_expr_relational_and_arm_band
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var n=85; __Check((n switch{>=80 and <90=>"B",_=>"other"}).ToString(), "B");
