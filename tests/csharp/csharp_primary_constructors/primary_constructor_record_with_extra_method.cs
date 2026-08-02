// vybe-test: csharp/csharp_primary_constructors/primary_constructor_record_with_extra_method
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y) { public int Sum() => X + Y; }
__Check((new Point(2, 3).Sum()).ToString(), "5");
