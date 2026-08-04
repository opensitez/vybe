// vybe-test: csharp/csharp_pattern_property/inheritance_property_pattern_on_derived_type
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

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

class Animal { public string Kind; } class Dog : Animal { public int Legs; } object o=new Dog{Kind="pet",Legs=4}; __P((o is Dog{Legs:4,Kind:"pet"}).ToString());
__Check("True");
