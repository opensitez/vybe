// vybe-test: csharp/csharp_oop_polymorphism/method_hiding_with_new_does_not_override_base_dispatch
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base{public virtual string Speak()=>"base";}
class Derived:Base{public new string Speak()=>"hidden";}
Base obj=new Derived();
__Check((obj.Speak()).ToString(), "base");
