// vybe-test: csharp/csharp_classes/class_basic
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((p.Describe()).ToString(), "Alice is 30");
