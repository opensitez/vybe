// vybe-test: csharp/csharp_collection_initializer_syntax/nested_object_initializer_builds_graph_in_one_expression
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Address { public string City { get; set; } }
class Person { public Address Home { get; set; } }
var person = new Person { Home = new Address { City = "Oslo" } };
__Check((person.Home.City).ToString(), "Oslo");
