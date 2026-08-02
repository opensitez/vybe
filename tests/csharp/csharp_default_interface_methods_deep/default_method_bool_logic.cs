// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_bool_logic
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFlag{bool On{get;} bool IsOff()=>!On;} class Switch:IFlag{public bool On{get;set;}=true;} __Check((new Switch().IsOff()).ToString(), "False");
