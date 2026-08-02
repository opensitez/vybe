// vybe-test: csharp/oop_advanced/readonly_auto_property
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Person {
    public string Name { get; }
    public Person(string name) { Name = name; }
}
var p = new Person("Alice");
__Check((p.Name).ToString(), "Alice");
