// vybe-test: csharp/csharp_oop_polymorphism/is_operator_succeeds_for_derived_held_as_base
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

class Animal{} class Dog:Animal{}
Animal a=new Dog();
__P((a is Dog).ToString()); __P((a is Animal).ToString());
__Check("True\nTrue");
