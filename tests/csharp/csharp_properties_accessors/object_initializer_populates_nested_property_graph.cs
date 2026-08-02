// vybe-test: csharp/csharp_properties_accessors/object_initializer_populates_nested_property_graph
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Address {
    public string City { get; set; }
}
class Office {
    public string Name { get; set; }
    public Address Location { get; set; }
}
var office = new Office {
    Name = "HQ",
    Location = new Address { City = "Paris" }
};
__Check((office.Name).ToString(), "HQ");
__Check((office.Location.City).ToString(), "Paris");
