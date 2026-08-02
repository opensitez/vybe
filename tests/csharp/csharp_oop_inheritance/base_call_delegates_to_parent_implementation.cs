// vybe-test: csharp/csharp_oop_inheritance/base_call_delegates_to_parent_implementation
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A { public virtual string Greet() => "Hello"; }
class B : A { public override string Greet() => base.Greet() + " World"; }
__Check((new B().Greet()).ToString(), "Hello World");
