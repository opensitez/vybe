// vybe-test: csharp/csharp_object_initializers/nested_object_initializer_sets_inner_object
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Address{public string City;}
class Person{public string Name;public Address Home;}
var p=new Person{Name="Bob",Home=new Address{City="Paris"}};
__Check((p.Home.City).ToString(), "Paris");
