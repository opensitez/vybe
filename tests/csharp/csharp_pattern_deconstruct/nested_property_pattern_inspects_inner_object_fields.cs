// vybe-test: csharp/csharp_pattern_deconstruct/nested_property_pattern_inspects_inner_object_fields
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

class Address { public string City; }
class Person { public Address Home; }
object p = new Person { Home = new Address { City = "Paris" } };
if (p is Person { Home: { City: "Paris" } }) Console.WriteLine("Paris");
else Console.WriteLine("elsewhere");
