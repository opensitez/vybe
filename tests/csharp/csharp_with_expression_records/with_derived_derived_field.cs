// vybe-test: csharp/csharp_with_expression_records/with_derived_derived_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Animal(string Name); record Dog(string Name,string Breed):Animal(Name); var k=(new Dog("Rex","Lab")) with{Breed="Pug"}; __Check((k.Breed).ToString(), "Pug");
