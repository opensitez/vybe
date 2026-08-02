// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_inheritance
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Pets;
class Animal { public string Name = "base"; }
class Dog : Animal { public string Breed = "lab"; }
var d = new Dog();
__Check((d.Name).ToString(), "base"); __Check((d.Breed).ToString(), "lab");
