// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_indexer_helper
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IIdx{string this[int i]{get;} string First()=>this[0];} class Arr:IIdx{public string this[int i]=>"v"+i;} __Check((new Arr().First()).ToString(), "v0");
