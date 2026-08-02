// vybe-test: csharp/csharp_primary_constructors/primary_constructor_double_param_formatted
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Rate(double value) { public double Value => value; }
__Check((new Rate(2.5).Value).ToString(), "2.5");
