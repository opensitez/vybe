// vybe-test: csharp/csharp_null_propagation/nested_null_conditional_walks_through_property_chain
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

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

class Address { public string City { get; set; } } class User { public Address Address { get; set; } } var user = new User { Address = new Address { City = "Paris" } }; __P((user?.Address?.City ?? "none").ToString());
__Check("Paris");
