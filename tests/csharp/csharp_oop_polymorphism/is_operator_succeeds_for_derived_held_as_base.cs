// vybe-test: csharp/csharp_oop_polymorphism/is_operator_succeeds_for_derived_held_as_base
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal{} class Dog:Animal{}
Animal a=new Dog();
__Check((a is Dog).ToString(), "True"); __Check((a is Animal).ToString(), "True");
