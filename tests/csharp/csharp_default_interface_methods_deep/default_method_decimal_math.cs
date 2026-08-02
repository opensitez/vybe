// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_decimal_math
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IMoney{decimal Add(decimal a,decimal b)=>a+b;} class Wallet:IMoney{} __Check((new Wallet().Add(1.5m,2.5m)).ToString(), "4.0");
