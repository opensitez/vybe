// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_generic_self_type_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ISelf<T> where T:ISelf<T>{static abstract T Identity(T v);}
struct Wrap:ISelf<Wrap>{public int N; public static Wrap Identity(Wrap v)=>v;}
var w=new Wrap{N=3}; __Check((Wrap.Identity(w).N).ToString(), "3");
