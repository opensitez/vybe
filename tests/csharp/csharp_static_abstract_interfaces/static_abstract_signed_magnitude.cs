// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_signed_magnitude
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ISign<T> where T:ISign<T>{static abstract T Negate(T v); static abstract T Abs(T v);}
struct IntWrap:ISign<IntWrap>{public int N; public static IntWrap Negate(IntWrap v)=>new IntWrap{N=-v.N}; public static IntWrap Abs(IntWrap v)=>new IntWrap{N=v.N<0?-v.N:v.N};}
__Check((IntWrap.Abs(new IntWrap{N=-4}).N).ToString(), "4");
