// vybe-test: csharp/csharp_primary_constructors/primary_constructor_derived_passes_param_to_base
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

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

class Animal(string name) { public string Name => name; }
class Dog(string name, string breed) : Animal(name) { public string Breed => breed; }
var d = new Dog("Rex", "Lab");
__P((d.Name).ToString()); __P((d.Breed).ToString());
__Check("Rex\nLab");
