// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_nullable_int_handling
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface INull{int? N{get;} int OrZero()=>N??0;} class Maybe:INull{public int? N{get;set;}=null;} __Check((new Maybe().OrZero()).ToString(), "0");
