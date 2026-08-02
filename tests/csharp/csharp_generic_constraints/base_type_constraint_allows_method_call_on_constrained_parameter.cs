// vybe-test: csharp/csharp_generic_constraints/base_type_constraint_allows_method_call_on_constrained_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal { public virtual string Sound() => "..."; }
class Dog : Animal { public override string Sound() => "woof"; }
string Speak<T>(T t) where T : Animal => t.Sound();
__Check((Speak(new Dog())).ToString(), "woof");
