// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_with_local_variable
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ILocal{int Triple(int n){var t=n*3; return t;}} class L:ILocal{} __Check((new L().Triple(4)).ToString(), "12");
