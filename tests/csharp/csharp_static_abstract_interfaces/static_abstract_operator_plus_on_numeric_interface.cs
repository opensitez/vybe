// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_operator_plus_on_numeric_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IAddable<T> where T:IAddable<T>{static abstract T operator+(T a,T b);}
struct Vec:IAddable<Vec>{public int X; public static Vec operator+(Vec a,Vec b)=>new Vec{X=a.X+b.X};}
__Check(((new Vec{X=2}+new Vec{X=3}).X).ToString(), "5");
