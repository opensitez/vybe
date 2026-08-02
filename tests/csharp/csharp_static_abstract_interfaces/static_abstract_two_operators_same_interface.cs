// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_two_operators_same_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IOps<T> where T:IOps<T>{static abstract T operator+(T a,T b); static abstract T operator*(T a,int k);}
struct Scale:IOps<Scale>{public int V; public static Scale operator+(Scale a,Scale b)=>new Scale{V=a.V+b.V}; public static Scale operator*(Scale a,int k)=>new Scale{V=a.V*k};}
__Check(((new Scale{V=2}*3).V).ToString(), "6");
