// vybe-test: csharp/csharp_typeof_vs_gettype/typeof_reports_declared_type_while_gettype_reports_runtime_type_of_instance
// origin: languages/csharp/tests/csharp/test_csharp_typeof_vs_gettype.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal { }
class Dog : Animal { }
Animal pet = new Dog();
__Check((typeof(Animal).Name).ToString(), "Animal");
__Check((pet.GetType().Name).ToString(), "Dog");
