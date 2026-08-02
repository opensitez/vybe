// vybe-test: csharp/csharp_switch_expression_core/switch_expr_enum_symbolic_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Mode { Off, On } var m=Mode.On; __Check((m switch{Mode.Off=>"0",Mode.On=>"1",_=>"?"}).ToString(), "1");
