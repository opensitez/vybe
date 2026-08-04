// vybe-test: csharp/csharp_oop_advanced2/covariant_return_type_narrows_return_of_override
// origin: languages/csharp/tests/csharp/test_csharp_oop_advanced2.rs

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

class Base{public virtual object Create()=>new object();}
class Derived:Base{public override string Create()=>"derived";}
Derived d=new Derived();
__P((d.Create()).ToString());
__Check("derived");
