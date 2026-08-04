// vybe-test: csharp/csharp_properties_accessors/object_initializer_populates_nested_property_graph
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

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
__P((office.Name).ToString());
__P((office.Location.City).ToString());
__Check("HQ\nParis");
