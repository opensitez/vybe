// vybe-test: csharp/csharp_primary_constructors/primary_constructor_char_param_returned
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Symbol(char ch) { public char Value => ch; }
__Check((new Symbol('Q').Value).ToString(), "Q");
