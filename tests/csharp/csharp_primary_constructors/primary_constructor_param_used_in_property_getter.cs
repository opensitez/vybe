// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_used_in_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Radius(int value) { public int Value => value; }
__Check((new Radius(7).Value).ToString(), "7");
