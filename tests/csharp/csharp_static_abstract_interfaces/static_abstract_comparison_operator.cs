// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_comparison_operator
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IComparableStatic<T> where T:IComparableStatic<T>{static abstract bool operator<(T a,T b);}
struct Rank:IComparableStatic<Rank>{public int Level; public static bool operator<(Rank a,Rank b)=>a.Level<b.Level;}
__Check((new Rank{Level=1}<new Rank{Level=2}).ToString(), "True");
