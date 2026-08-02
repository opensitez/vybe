// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_property_getter_on_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IUnit<T> where T:IUnit<T>{static abstract T Zero{get;}}
struct Counter:IUnit<Counter>{public int V; public static Counter Zero=>new Counter{V=0};}
__Check((Counter.Zero.V).ToString(), "0");
