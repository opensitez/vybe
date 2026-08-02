// vybe-test: csharp/csharp_default_interface_methods_deep/diamond_resolution_prefers_most_specific_class_method
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IA{string Tag()=>"A";} interface IB:IA{string Tag()=>"B";} class Leaf:IB{public string Tag()=>"L";} __Check((new Leaf().Tag()).ToString(), "L");
