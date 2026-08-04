// vybe-test: csharp/csharp_classes/class_basic
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

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

class Person {
    public string Name;
    public int Age;
    public Person(string name, int age) {
        Name = name;
        Age = age;
    }
    public string Describe() {
        return Name + " is " + Age;
    }
}
var p = new Person("Alice", 30);
__P((p.Describe()).ToString());
__Check("Alice is 30");
