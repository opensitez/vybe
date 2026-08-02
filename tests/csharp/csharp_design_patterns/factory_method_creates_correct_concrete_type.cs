// vybe-test: csharp/csharp_design_patterns/factory_method_creates_correct_concrete_type
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Animal{public abstract string Sound();}
class Dog:Animal{public override string Sound()=>"woof";}
class Cat:Animal{public override string Sound()=>"meow";}
Animal Create(string kind)=>kind=="dog"?(Animal)new Dog():new Cat();
__Check((Create("dog").Sound()).ToString(), "woof");
__Check((Create("cat").Sound()).ToString(), "meow");
