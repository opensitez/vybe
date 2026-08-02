// vybe-test: csharp/csharp_const_and_readonly_fields/readonly_field_cannot_change_after_constructor_body_completes
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
    public readonly int Seed;
    public Counter(int seed) { Seed = seed; }
    public int Read() { return Seed; }
}
__Check((new Counter(3).Read()).ToString(), "3");
