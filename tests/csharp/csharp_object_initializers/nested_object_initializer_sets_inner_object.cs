// vybe-test: csharp/csharp_object_initializers/nested_object_initializer_sets_inner_object
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

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

class Address{public string City;}
class Person{public string Name;public Address Home;}
var p=new Person{Name="Bob",Home=new Address{City="Paris"}};
__P((p.Home.City).ToString());
__Check("Paris");
