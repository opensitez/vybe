// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_struct_implementor
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IArea{double Area()=>1.0;} struct Unit:IArea{} __Check((new Unit().Area()).ToString(), "1");
