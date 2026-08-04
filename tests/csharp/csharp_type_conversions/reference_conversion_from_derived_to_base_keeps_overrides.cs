// vybe-test: csharp/csharp_type_conversions/reference_conversion_from_derived_to_base_keeps_overrides
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

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

class Animal { public virtual string Speak() { return "animal"; } } class Dog : Animal { public override string Speak() { return "woof"; } } Dog dog = new Dog(); Animal animal = dog; __P((animal.Speak()).ToString());
__Check("woof");
