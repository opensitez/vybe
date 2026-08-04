// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_property_default_on_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

interface IDefault<T> where T:IDefault<T>{static virtual T Fallback=>default; static abstract T Primary();}
struct Item:IDefault<Item>{public int Id; public static Item Primary()=>new Item{Id=1}; public static Item Fallback=>new Item{Id=99};}
__P((Item.Primary().Id).ToString());
__Check("1");
