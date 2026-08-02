// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_derived_record
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Animal(string Name); record Dog(string Name,int Age):Animal(Name); var (name,age)=new Dog("Rex",4); __Check((name).ToString(), "Rex"); __Check((age).ToString(), "4");
