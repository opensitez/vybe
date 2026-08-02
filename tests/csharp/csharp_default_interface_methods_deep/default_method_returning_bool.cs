// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_returning_bool
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ICheck{bool Ok()=>true;} class Gate:ICheck{} __Check((new Gate().Ok()).ToString(), "True");
