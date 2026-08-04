// vybe-test: csharp/csharp_oop_advanced2/sealed_override_prevents_further_overriding
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

class A{public virtual string Tag()=>"A";}
class B:A{public sealed override string Tag()=>"B";}
class C:B{}
C c=new C();
__P((c.Tag()).ToString());
__Check("B");
