// vybe-test: csharp/csharp_properties_accessors/object_initializer_populates_nested_property_graph
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

using static __Harness;

var office = new Office {
    Name = "HQ",
    Location = new Address { City = "Paris" }
}
;
__P((office.Name).ToString());
__P((office.Location.City).ToString());
__Check("HQ\nParis");

class Address {
    public string City { get; set; }
}

class Office {
    public string Name { get; set; }
    public Address Location { get; set; }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
