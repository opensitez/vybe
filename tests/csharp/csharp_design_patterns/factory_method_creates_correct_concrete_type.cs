// vybe-test: csharp/csharp_design_patterns/factory_method_creates_correct_concrete_type
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

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

abstract class Animal{public abstract string Sound();}
class Dog:Animal{public override string Sound()=>"woof";}
class Cat:Animal{public override string Sound()=>"meow";}
Animal Create(string kind)=>kind=="dog"?(Animal)new Dog():new Cat();
__P((Create("dog").Sound()).ToString());
__P((Create("cat").Sound()).ToString());
__Check("woof\nmeow");
