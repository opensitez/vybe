// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_reused_by_two_classes
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IDouble{int Twice(int n)=>n*2;} class A:IDouble{} class B:IDouble{} __Check((new A().Twice(3)+new B().Twice(4)).ToString(), "14");
