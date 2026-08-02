// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_on_generic_interface
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBox<T>{T Echo(T v)=>v;} class IntBox:IBox<int>{} __Check((new IntBox().Echo(7)).ToString(), "7");
