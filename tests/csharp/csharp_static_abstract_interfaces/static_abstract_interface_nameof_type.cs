// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_interface_nameof_type
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IName<T> where T:IName<T>{static abstract string TypeName();}
struct Named:IName<Named>{public static string TypeName()=>nameof(Named);}
__Check((Named.TypeName()).ToString(), "Named");
