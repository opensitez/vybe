// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_override_calls_base_default_via_super_not_available_use_explicit
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IA{string V()=>"a";} interface IB:IA{string W()=>V()+"b";} class Z:IB{public string V()=>"z";} __Check((((IB)new Z()).W()).ToString(), "ab");
