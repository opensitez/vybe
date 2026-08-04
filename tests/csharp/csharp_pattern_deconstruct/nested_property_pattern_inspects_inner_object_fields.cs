// vybe-test: csharp/csharp_pattern_deconstruct/nested_property_pattern_inspects_inner_object_fields
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

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

class Address { public string City; }
class Person { public Address Home; }
object p = new Person { Home = new Address { City = "Paris" } };
if (p is Person { Home: { City: "Paris" } }) __P(("Paris").ToString());
else __P(("elsewhere").ToString());
__Check("Paris");
