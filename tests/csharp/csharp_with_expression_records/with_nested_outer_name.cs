// vybe-test: csharp/csharp_with_expression_records/with_nested_outer_name
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Address(string City); record Person(string Name,Address Home); var q=(new Person("Ann",new Address("Oslo"))) with{Name="Bob"}; __Check((q.Name).ToString(), "Bob");
