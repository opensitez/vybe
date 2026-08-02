// vybe-test: csharp/csharp_default_interface_methods_deep/default_interface_method_called_on_concrete_without_override
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IGreet{string Hello()=>"hi";} class Person:IGreet{} __Check((new Person().Hello()).ToString(), "hi");
