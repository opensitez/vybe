// vybe-test: csharp/csharp_records_advanced/record_inheritance_preserves_base_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Animal(string Name); record Dog(string Name, int Age) : Animal(Name); var dog = new Dog("Rex", 5); __Check((dog.Name).ToString(), "Rex"); __Check((dog.Age).ToString(), "5");
