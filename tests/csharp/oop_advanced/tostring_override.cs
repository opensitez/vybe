// vybe-test: csharp/oop_advanced/tostring_override
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Person {
    public string Name { get; set; }
    public int Age { get; set; }
    public Person(string name, int age) { Name = name; Age = age; }
    public override string ToString() { return Name + " (" + Age + ")"; }
}
var p = new Person("Alice", 30);
__Check((p.ToString()).ToString(), "Alice (30)");
__Check((p).ToString(), "Alice (30)");
