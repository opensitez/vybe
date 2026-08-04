// vybe-test: csharp/classes/class_with_constructor
// origin: languages/csharp/tests/csharp/test_classes.rs

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
            string name;
            int age;
            public Person(string n, int a) {
                this.name = n;
                this.age = a;
            }
            public string Describe() {
                return this.name + " is " + this.age;
            }
        }
        var p = new Person("Alice", 30);
        __P((p.Describe()).ToString());
__Check("Alice is 30");
