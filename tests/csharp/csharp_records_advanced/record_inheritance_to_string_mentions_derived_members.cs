// vybe-test: csharp/csharp_records_advanced/record_inheritance_to_string_mentions_derived_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Animal(string Name); record Cat(string Name, string Color) : Animal(Name); __Check((new Cat("Milo", "Black").ToString().Contains("Color = Black")).ToString(), "True");
