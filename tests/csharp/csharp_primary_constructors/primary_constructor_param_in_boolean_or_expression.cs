// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_in_boolean_or_expression
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Flags(bool a, bool b) { public bool Any => a || b; }
__Check((new Flags(false, true).Any).ToString(), "True");
