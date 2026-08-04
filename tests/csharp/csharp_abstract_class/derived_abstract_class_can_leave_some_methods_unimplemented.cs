// vybe-test: csharp/csharp_abstract_class/derived_abstract_class_can_leave_some_methods_unimplemented
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

abstract class A{public abstract int X();public abstract int Y();}
abstract class B:A{public override int X()=>1;}
class C:B{public override int Y()=>2;}
var c=new C();
__P((c.X()).ToString()); __P((c.Y()).ToString());
__Check("1\n2");
