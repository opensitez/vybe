// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_uses_instance_property
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IHas{int N{get;} int Twice()=>N*2;} class Box:IHas{public int N{get;set;}} var b=new Box{N=5}; __Check((b.Twice()).ToString(), "10");
