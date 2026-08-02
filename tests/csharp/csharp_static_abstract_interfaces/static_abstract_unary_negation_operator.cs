// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_unary_negation_operator
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface INegatable<T> where T:INegatable<T>{static abstract T operator-(T v);}
struct Signed:INegatable<Signed>{public int N; public static Signed operator-(Signed v)=>new Signed{N=-v.N};}
__Check(((-new Signed{N=8}).N).ToString(), "-8");
