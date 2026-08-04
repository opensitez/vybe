// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_three_string_fields
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

class Street { public string Name; } class Addr { public Street S; } class Person { public Addr A; } object p=new Person{A=new Addr{S=new Street{Name="Main"}}}; __P((p is Person{A:{S:{Name:"Main"}}}).ToString());
__Check("True");
