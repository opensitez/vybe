// vybe-test: csharp/csharp_abstract_class/abstract_class_with_constructor_initialized_by_derived
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Named{public string Name;public Named(string n){Name=n;}}
class Tag:Named{public Tag(string n):base(n){}}
__Check((new Tag("admin").Name).ToString(), "admin");
