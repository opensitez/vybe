// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_string_on_inner
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Address { public string City; } class Person { public Address Home; } object p=new Person{Home=new Address{City="Paris"}}; __Check((p is Person{Home:{City:"Paris"}}).ToString(), "True");
