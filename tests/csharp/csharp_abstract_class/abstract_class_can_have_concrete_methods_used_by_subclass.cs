// vybe-test: csharp/csharp_abstract_class/abstract_class_can_have_concrete_methods_used_by_subclass
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

abstract class Animal{
    public abstract string Sound();
    public string Speak()=>$"I say {Sound()}";
}
class Cat:Animal{public override string Sound()=>"meow";}
__P((new Cat().Speak()).ToString());
__Check("I say meow");
