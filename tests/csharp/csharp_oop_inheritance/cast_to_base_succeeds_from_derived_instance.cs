// vybe-test: csharp/csharp_oop_inheritance/cast_to_base_succeeds_from_derived_instance
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { public int X = 1; }
class Derived : Base { public int Y = 2; }
Base b = (Base)new Derived();
__Check((b.X).ToString(), "1");
