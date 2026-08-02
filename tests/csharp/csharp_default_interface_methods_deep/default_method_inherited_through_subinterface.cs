// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_inherited_through_subinterface
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IRoot{int Base()=>1;} interface IChild:IRoot{int Child()=>Base()+1;} class Node:IChild{} __Check((new Node().Child()).ToString(), "2");
