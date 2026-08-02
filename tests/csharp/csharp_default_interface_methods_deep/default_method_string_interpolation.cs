// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_string_interpolation
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IName{string Name{get;} string Label()=>$"name={Name}";} class User:IName{public string Name{get;set;}="Ann";} __Check((new User().Label()).ToString(), "name=Ann");
