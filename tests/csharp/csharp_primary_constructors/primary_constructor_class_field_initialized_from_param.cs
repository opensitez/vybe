// vybe-test: csharp/csharp_primary_constructors/primary_constructor_class_field_initialized_from_param
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Holder(int seed) { int value = seed; public int Read() => value; }
__Check((new Holder(99).Read()).ToString(), "99");
