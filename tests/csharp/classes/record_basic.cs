// vybe-test: csharp/classes/record_basic
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Person(string Name, int Age);
        var p = new Person("Alice", 30);
        __Check((p.Name).ToString(), "Alice");
        __Check((p.Age).ToString(), "30");
