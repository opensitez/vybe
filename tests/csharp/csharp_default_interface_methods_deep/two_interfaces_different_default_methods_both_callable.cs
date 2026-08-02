// vybe-test: csharp/csharp_default_interface_methods_deep/two_interfaces_different_default_methods_both_callable
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IA{int A()=>1;} interface IB{int B()=>2;} class Both:IA,IB{} var x=new Both(); __Check((x.A()+x.B()).ToString(), "3");
