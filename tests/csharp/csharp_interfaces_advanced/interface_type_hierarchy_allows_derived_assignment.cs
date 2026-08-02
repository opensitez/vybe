// vybe-test: csharp/csharp_interfaces_advanced/interface_type_hierarchy_allows_derived_assignment
// origin: languages/csharp/tests/csharp/test_csharp_interfaces_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IAnimal{string Kind();}
interface IPet:IAnimal{string Name();}
class Dog:IPet{public string Kind()=>"dog"; public string Name()=>"Rex";}
IPet pet=new Dog();
IAnimal animal=pet;
__Check((animal.Kind()).ToString(), "dog");
__Check((pet.Name()).ToString(), "Rex");
