// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_multiple_interface_inheritance_same_signature
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IA{int Score()=>1;} interface IB{int Score()=>2;} class Dual:IA,IB{public int Score()=>3;} __Check((new Dual().Score()).ToString(), "3");
