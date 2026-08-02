// vybe-test: csharp/csharp_properties_accessors/init_only_property_is_set_by_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Customer {
    public string Name { get; init; }
    public int Tier { get; init; }
}
var customer = new Customer { Name = "Ada", Tier = 3 };
__Check((customer.Name).ToString(), "Ada");
__Check((customer.Tier).ToString(), "3");
