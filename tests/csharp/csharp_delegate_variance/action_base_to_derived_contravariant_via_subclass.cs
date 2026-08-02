// vybe-test: csharp/csharp_delegate_variance/action_base_to_derived_contravariant_via_subclass
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal{} class Dog:Animal{} System.Action<Animal> feed=a=>__Check((a!=null).ToString(), "True"); System.Action<Dog> feedDog=feed; feedDog(new Dog());
