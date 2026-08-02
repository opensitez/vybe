// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_equality_between_instances
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Id(int Value);
var a = new Id(5);
var b = new Id(5);
__Check((a == b).ToString(), "True");
