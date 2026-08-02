// vybe-test: csharp/csharp_abstract_class/abstract_class_can_have_concrete_methods_used_by_subclass
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Animal{
    public abstract string Sound();
    public string Speak()=>$"I say {Sound()}";
}
class Cat:Animal{public override string Sound()=>"meow";}
__Check((new Cat().Speak()).ToString(), "I say meow");
