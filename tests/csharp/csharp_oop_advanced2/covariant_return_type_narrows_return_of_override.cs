// vybe-test: csharp/csharp_oop_advanced2/covariant_return_type_narrows_return_of_override
// origin: languages/csharp/tests/csharp/test_csharp_oop_advanced2.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base{public virtual object Create()=>new object();}
class Derived:Base{public override string Create()=>"derived";}
Derived d=new Derived();
__Check((d.Create()).ToString(), "derived");
