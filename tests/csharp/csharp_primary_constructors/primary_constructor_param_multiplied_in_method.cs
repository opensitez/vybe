// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_multiplied_in_method
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Scale(int factor) { public int Apply(int n) => n * factor; }
__Check((new Scale(5).Apply(6)).ToString(), "30");
