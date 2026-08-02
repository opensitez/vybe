// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_with_parameters
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IAdd{int Sum(int a,int b)=>a+b;} class Calc:IAdd{} __Check((new Calc().Sum(2,5)).ToString(), "7");
