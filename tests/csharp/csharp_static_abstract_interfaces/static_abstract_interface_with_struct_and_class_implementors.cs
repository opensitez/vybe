// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_interface_with_struct_and_class_implementors
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IShared<T> where T:IShared<T>{static abstract int Key();}
struct SA:IShared<SA>{public static int Key()=>1;} class CA:IShared<CA>{public static int Key()=>2;}
__Check((SA.Key()+CA.Key()).ToString(), "3");
