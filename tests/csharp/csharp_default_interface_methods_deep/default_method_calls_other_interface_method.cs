// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_calls_other_interface_method
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBase{string Core()=>"core"; string Wrap()=>"["+Core()+"]";} class Node:IBase{} __Check((new Node().Wrap()).ToString(), "[core]");
