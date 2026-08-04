// vybe-test: csharp/csharp_typeof_vs_gettype/typeof_reports_declared_type_while_gettype_reports_runtime_type_of_instance
// origin: languages/csharp/tests/csharp/test_csharp_typeof_vs_gettype.rs

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

class Animal { }
class Dog : Animal { }
Animal pet = new Dog();
__P((typeof(Animal).Name).ToString());
__P((pet.GetType().Name).ToString());
__Check("Animal\nDog");
