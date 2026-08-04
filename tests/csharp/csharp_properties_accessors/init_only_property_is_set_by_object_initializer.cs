// vybe-test: csharp/csharp_properties_accessors/init_only_property_is_set_by_object_initializer
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

class Customer {
    public string Name { get; init; }
    public int Tier { get; init; }
}
var customer = new Customer { Name = "Ada", Tier = 3 };
__P((customer.Name).ToString());
__P((customer.Tier).ToString());
__Check("Ada\n3");
