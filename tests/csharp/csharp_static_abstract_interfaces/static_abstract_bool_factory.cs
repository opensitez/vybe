// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_bool_factory
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFlag<T> where T:IFlag<T>{static abstract T True(); static abstract T False();}
struct Bit:IFlag<Bit>{public bool On; public static Bit True()=>new Bit{On=true}; public static Bit False()=>new Bit{On=false};}
__Check((Bit.True().On).ToString(), "True");
