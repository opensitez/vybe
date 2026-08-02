// vybe-test: csharp/csharp_constructor_patterns/primary_constructor_on_record_sets_all_fields
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Person(string Name,int Age);
var p=new Person("Grace",40);
__Check((p.Name).ToString(), "Grace"); __Check((p.Age).ToString(), "40");
