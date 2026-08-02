// vybe-test: csharp/csharp_primary_constructors/primary_constructor_nested_type_on_outer
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer(int seed) {
    public class Inner { public int Value; }
    public Inner Make() => new Inner { Value = seed };
}
__Check((new Outer(6).Make().Value).ToString(), "6");
