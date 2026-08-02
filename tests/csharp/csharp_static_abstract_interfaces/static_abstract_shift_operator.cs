// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_shift_operator
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IShift<T> where T:IShift<T>{static abstract T operator<<(T v,int n);}
struct Bits:IShift<Bits>{public int V; public static Bits operator<<(Bits v,int n)=>new Bits{V=v.V<<n};}
__Check(((new Bits{V=1}<<3).V).ToString(), "8");
