// vybe-test: csharp/csharp_primary_constructors/primary_constructor_decimal_param_stored
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Money(decimal amount) { public decimal Amount => amount; }
__Check((new Money(9.99m).Amount).ToString(), "9.99");
