// vybe-test: csharp/csharp_oop_advanced2/multiple_levels_of_base_call_chain
// origin: languages/csharp/tests/csharp/test_csharp_oop_advanced2.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A{public virtual string Name()=>"A";}
class B:A{public override string Name()=>"B+"+base.Name();}
class C:B{public override string Name()=>"C+"+base.Name();}
__Check((new C().Name()).ToString(), "C+B+A");
