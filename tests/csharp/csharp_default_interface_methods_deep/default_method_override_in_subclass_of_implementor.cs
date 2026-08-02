// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_override_in_subclass_of_implementor
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IVal{int Get()=>0;} class Base:IVal{} class Derived:Base,IVal{public int Get()=>5;} __Check((new Derived().Get()).ToString(), "5");
