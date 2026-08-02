// vybe-test: csharp/csharp_oop_advanced2/sealed_override_prevents_further_overriding
// origin: languages/csharp/tests/csharp/test_csharp_oop_advanced2.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A{public virtual string Tag()=>"A";}
class B:A{public sealed override string Tag()=>"B";}
class C:B{}
C c=new C();
__Check((c.Tag()).ToString(), "B");
