// vybe-test: csharp/csharp_object_initializers/object_initializer_sets_multiple_properties
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Person{public string Name;public int Age;}
var p=new Person{Name="Alice",Age=30};
__Check((p.Name).ToString(), "Alice"); __Check((p.Age).ToString(), "30");
