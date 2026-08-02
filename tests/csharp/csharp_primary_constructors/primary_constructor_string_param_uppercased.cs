// vybe-test: csharp/csharp_primary_constructors/primary_constructor_string_param_uppercased
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Shout(string word) { public string Loud() => word.ToUpper(); }
__Check((new Shout("go").Loud()).ToString(), "GO");
