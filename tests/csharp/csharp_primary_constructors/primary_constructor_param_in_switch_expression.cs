// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_in_switch_expression
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Mode(int code) { public string Name() => code switch { 1 => "a", 2 => "b", _ => "x" }; }
__Check((new Mode(2).Name()).ToString(), "b");
