// vybe-test: csharp/csharp_with_expression_records/with_derived_base_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Animal(string Name); record Dog(string Name,string Breed):Animal(Name); var k=(new Dog("Rex","Lab")) with{Name="Max"}; __Check((k.Name).ToString(), "Max");
