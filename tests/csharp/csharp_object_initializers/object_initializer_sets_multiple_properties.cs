// vybe-test: csharp/csharp_object_initializers/object_initializer_sets_multiple_properties
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

class Person{public string Name;public int Age;}
var p=new Person{Name="Alice",Age=30};
__P((p.Name).ToString()); __P((p.Age).ToString());
__Check("Alice\n30");
