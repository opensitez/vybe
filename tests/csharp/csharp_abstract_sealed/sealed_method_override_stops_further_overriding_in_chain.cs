// vybe-test: csharp/csharp_abstract_sealed/sealed_method_override_stops_further_overriding_in_chain
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A { public virtual string Name() => "A"; }
class B : A { public sealed override string Name() => "B"; }
class C : B { }
A obj = new C();
__Check((obj.Name()).ToString(), "B");
