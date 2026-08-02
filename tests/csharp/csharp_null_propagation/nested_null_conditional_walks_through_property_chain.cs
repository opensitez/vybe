// vybe-test: csharp/csharp_null_propagation/nested_null_conditional_walks_through_property_chain
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Address { public string City { get; set; } } class User { public Address Address { get; set; } } var user = new User { Address = new Address { City = "Paris" } }; __Check((user?.Address?.City ?? "none").ToString(), "Paris");
