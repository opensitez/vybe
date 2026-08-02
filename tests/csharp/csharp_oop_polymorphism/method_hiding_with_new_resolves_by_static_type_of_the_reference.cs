// vybe-test: csharp/csharp_oop_polymorphism/method_hiding_with_new_resolves_by_static_type_of_the_reference
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base{public virtual string Speak()=>"base";}
class Derived:Base{public new string Speak()=>"hidden";}
Derived d=new Derived();
Base b=d;
__Check((d.Speak()).ToString(), "hidden");
__Check((b.Speak()).ToString(), "base");
