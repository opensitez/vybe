// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_with_null_conditional
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IMaybe{string Name{get;} string Safe()=>Name??"none";} class Item:IMaybe{public string Name{get;set;}} __Check((new Item().Safe()).ToString(), "none");
