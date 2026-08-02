// vybe-test: csharp/csharp_generics_where/where_base_class_constraint_calls_base_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Animal{public abstract string Sound();}
class Dog:Animal{public override string Sound()=>"woof";}
string Hear<T>(T a) where T:Animal=>a.Sound();
__Check((Hear(new Dog())).ToString(), "woof");
