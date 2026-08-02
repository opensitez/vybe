// vybe-test: csharp/csharp_with_expression_records/with_nested_replace_inner
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Address(string City); record Person(string Name,Address Home); var p=new Person("Ann",new Address("Oslo")); var q=p with{Home=new Address("Paris")}; __Check((q.Home.City).ToString(), "Paris");
