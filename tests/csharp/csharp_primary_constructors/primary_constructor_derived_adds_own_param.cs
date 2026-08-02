// vybe-test: csharp/csharp_primary_constructors/primary_constructor_derived_adds_own_param
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base(int x) { public int X => x; }
class Extra(int x, int y) : Base(x) { public int Y => y; }
__Check((new Extra(2, 5).Y).ToString(), "5");
