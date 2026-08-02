// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_bitwise_and_operator
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBit<T> where T:IBit<T>{static abstract T operator&(T a,T b);}
struct Mask:IBit<Mask>{public int Bits; public static Mask operator&(Mask a,Mask b)=>new Mask{Bits=a.Bits&b.Bits};}
__Check(((new Mask{Bits=7}&new Mask{Bits=3}).Bits).ToString(), "3");
