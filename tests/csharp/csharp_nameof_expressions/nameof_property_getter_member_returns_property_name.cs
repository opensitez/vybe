// vybe-test: csharp/csharp_nameof_expressions/nameof_property_getter_member_returns_property_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Person{public string Name{get;set;}} __Check((nameof(Person.Name)).ToString(), "Name");
