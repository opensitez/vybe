// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_method_default_calls_primary
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IDouble<T> where T:IDouble<T>{static abstract T One(); static virtual T Two(){return One();}}
struct Dup:IDouble<Dup>{public int V; public static Dup One()=>new Dup{V=1}; public static Dup Two()=>new Dup{V=2};}
__Check((Dup.Two().V).ToString(), "2");
