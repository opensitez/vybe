// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_rejects_wrong_inner
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

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

class Address { public string City; } class Person { public Address Home; } object p=new Person{Home=new Address{City="Paris"}}; __P((p is Person{Home:{City:"London"}}).ToString());
__Check("False");
