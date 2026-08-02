// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_replaced_only_on_one_interface_in_diamond
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IA{string Tag()=>"A";} interface IB{string Tag()=>"B";} class Pick:IA,IB{public string Tag()=>"P";} __Check((new Pick().Tag()).ToString(), "P");
