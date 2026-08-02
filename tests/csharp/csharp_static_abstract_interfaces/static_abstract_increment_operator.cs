// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_increment_operator
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IInc<T> where T:IInc<T>{static abstract T operator++(T v);}
struct Num:IInc<Num>{public int N; public static Num operator++(Num v)=>new Num{N=v.N+1};}
__Check(((++new Num{N=4}).N).ToString(), "5");
