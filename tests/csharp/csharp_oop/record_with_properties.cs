// vybe-test: csharp/csharp_oop/record_with_properties
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

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
