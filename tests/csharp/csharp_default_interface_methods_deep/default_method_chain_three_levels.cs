// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_chain_three_levels
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface I1{string S()=>"a";} interface I2:I1{string T()=>S()+"b";} class X:I2{} __Check((new X().T()).ToString(), "ab");
