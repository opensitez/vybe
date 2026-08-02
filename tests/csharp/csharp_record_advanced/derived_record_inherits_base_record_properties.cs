// vybe-test: csharp/csharp_record_advanced/derived_record_inherits_base_record_properties
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Animal(string Name);
record Dog(string Name,string Breed):Animal(Name);
var d=new Dog("Rex","Lab");
__Check((d.Name).ToString(), "Rex"); __Check((d.Breed).ToString(), "Lab");
