// vybe-test: csharp/csharp_oop_inheritance/derived_class_inherits_public_method_from_base
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { public string Hello() => "hello"; }
class Derived : Base { }
__Check((new Derived().Hello()).ToString(), "hello");
