// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_modulo_operator
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IMod<T> where T:IMod<T>{static abstract T operator%(T a,T b);}
struct Mod:IMod<Mod>{public int V; public static Mod operator%(Mod a,Mod b)=>new Mod{V=a.V%b.V};}
__Check(((new Mod{V=10}%new Mod{V=3}).V).ToString(), "1");
