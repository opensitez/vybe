// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_string_builder_pattern
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBuild{string Step1()=>"a"; string Step2()=>Step1()+"b";} class Chain:IBuild{} __Check((new Chain().Step2()).ToString(), "ab");
