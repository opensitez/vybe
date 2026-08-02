// vybe-test: csharp/csharp_oop_inheritance/abstract_class_provides_partial_implementation
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Base {
    public abstract int Value();
    public int Double() => Value() * 2;
}
class Impl : Base { public override int Value() => 5; }
__Check((new Impl().Double()).ToString(), "10");
