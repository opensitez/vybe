// vybe-test: csharp/csharp_oop_polymorphism/method_hiding_with_new_does_not_override_base_dispatch
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

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

class Base{public virtual string Speak()=>"base";}
class Derived:Base{public new string Speak()=>"hidden";}
Base obj=new Derived();
__P((obj.Speak()).ToString());
__Check("base");
