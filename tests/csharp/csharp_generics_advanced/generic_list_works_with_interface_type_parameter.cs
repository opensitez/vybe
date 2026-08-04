// vybe-test: csharp/csharp_generics_advanced/generic_list_works_with_interface_type_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

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

interface IAnimal { string Sound(); }
class Cat : IAnimal { public string Sound() => "meow"; }
var animals = new System.Collections.Generic.List<IAnimal> { new Cat() };
foreach(var a in animals) __P((a.Sound()).ToString());
__Check("meow");
