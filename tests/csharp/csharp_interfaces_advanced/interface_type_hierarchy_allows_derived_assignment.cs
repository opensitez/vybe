// vybe-test: csharp/csharp_interfaces_advanced/interface_type_hierarchy_allows_derived_assignment
// origin: languages/csharp/tests/csharp/test_csharp_interfaces_advanced.rs

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

interface IAnimal{string Kind();}
interface IPet:IAnimal{string Name();}
class Dog:IPet{public string Kind()=>"dog"; public string Name()=>"Rex";}
IPet pet=new Dog();
IAnimal animal=pet;
__P((animal.Kind()).ToString());
__P((pet.Name()).ToString());
__Check("dog\nRex");
