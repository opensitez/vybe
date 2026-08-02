// vybe-test: csharp/csharp_record_types/record_inheritance_shares_base_properties
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Animal(string Name);
record Dog(string Name, string Breed) : Animal(Name);
var d = new Dog("Rex","Lab");
__Check((d.Name).ToString(), "Rex"); __Check((d.Breed).ToString(), "Lab");
