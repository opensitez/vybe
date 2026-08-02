// vybe-test: csharp/modern_features/record_with_expression
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Person(string Name, int Age);
var p1 = new Person("Alice", 30);
var p2 = p1 with { Age = 31 };
__Check((p1).ToString(), "Person { Name = Alice, Age = 30 }");
__Check((p2).ToString(), "Person { Name = Alice, Age = 31 }");
