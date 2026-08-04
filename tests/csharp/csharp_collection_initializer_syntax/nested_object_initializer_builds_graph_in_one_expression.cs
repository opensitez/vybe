// vybe-test: csharp/csharp_collection_initializer_syntax/nested_object_initializer_builds_graph_in_one_expression
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

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

class Address { public string City { get; set; } }
class Person { public Address Home { get; set; } }
var person = new Person { Home = new Address { City = "Oslo" } };
__P((person.Home.City).ToString());
__Check("Oslo");
