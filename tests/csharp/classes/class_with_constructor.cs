// vybe-test: csharp/classes/class_with_constructor
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
        __Check((p.Describe()).ToString(), "Alice is 30");
