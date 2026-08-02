// vybe-test: csharp/csharp_oop_polymorphism/virtual_method_dispatches_to_most_derived_override
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base{public virtual string Speak()=>"base";}
class Derived:Base{public override string Speak()=>"derived";}
Base obj=new Derived();
__Check((obj.Speak()).ToString(), "derived");
