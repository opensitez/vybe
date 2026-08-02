// vybe-test: csharp/csharp_delegate_variance/func_derived_to_base_covariant_return_via_subclass
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal{} class Dog:Animal{} System.Func<Dog> getDog=()=>new Dog(); System.Func<Animal> getAnimal=getDog; __Check((getAnimal()!=null).ToString(), "True");
