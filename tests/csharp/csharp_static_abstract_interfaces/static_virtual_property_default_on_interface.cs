// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_property_default_on_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IDefault<T> where T:IDefault<T>{static virtual T Fallback=>default; static abstract T Primary();}
struct Item:IDefault<Item>{public int Id; public static Item Primary()=>new Item{Id=1}; public static Item Fallback=>new Item{Id=99};}
__Check((Item.Primary().Id).ToString(), "1");
