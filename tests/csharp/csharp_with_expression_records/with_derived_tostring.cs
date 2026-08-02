// vybe-test: csharp/csharp_with_expression_records/with_derived_tostring
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Animal(string Name); record Cat(string Name,string Color):Animal(Name); var d=(new Cat("M","W")) with{Color="B"}; __Check((d.ToString().Contains("B")).ToString(), "True");
