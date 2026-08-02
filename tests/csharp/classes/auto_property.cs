// vybe-test: csharp/classes/auto_property
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Person {
            public string Name { get; set; }
            public Person(string n) { this.Name = n; }
        }
        var p = new Person("Alice");
        __Check((p.Name).ToString(), "Alice");
