// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_to_string_in_method
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Code(int n) { public string Text() => n.ToString(); }
__Check((new Code(77).Text()).ToString(), "77");
