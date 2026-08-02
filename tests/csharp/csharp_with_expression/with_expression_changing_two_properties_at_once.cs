// vybe-test: csharp/csharp_with_expression/with_expression_changing_two_properties_at_once
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Person(string Name, int Age);
var p = new Person("Ada", 30);
var updated = p with { Name = "Grace", Age = 31 };
__Check((updated.Name).ToString(), "Grace");
__Check((updated.Age).ToString(), "31");
