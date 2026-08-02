// vybe-test: csharp/csharp_abstract_class/derived_abstract_class_can_leave_some_methods_unimplemented
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class A{public abstract int X();public abstract int Y();}
abstract class B:A{public override int X()=>1;}
class C:B{public override int Y()=>2;}
var c=new C();
__Check((c.X()).ToString(), "1"); __Check((c.Y()).ToString(), "2");
