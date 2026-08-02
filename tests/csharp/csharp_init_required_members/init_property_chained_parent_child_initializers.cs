// vybe-test: csharp/csharp_init_required_members/init_property_chained_parent_child_initializers
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Address { public string City { get; init; } }
class Person { public string Name { get; init; } public Address Home { get; init; } }
var p = new Person { Name = "Ann", Home = new Address { City = "Oslo" } };
__Check((p.Home.City).ToString(), "Oslo");
