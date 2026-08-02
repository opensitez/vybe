// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_two_defaults_only_one_overridden
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IA{int A()=>1;} interface IB{int B()=>2;} class Mix:IA,IB{public int A()=>10;} var m=new Mix(); __Check((m.A()+m.B()).ToString(), "12");
