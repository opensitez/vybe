// vybe-test: csharp/csharp_init_required_members/init_property_chained_parent_child_initializers
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

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

class Address { public string City { get; init; } }
class Person { public string Name { get; init; } public Address Home { get; init; } }
var p = new Person { Name = "Ann", Home = new Address { City = "Oslo" } };
__P((p.Home.City).ToString());
__Check("Oslo");
