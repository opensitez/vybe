// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_combined_add_and_parse
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IOps2<T> where T:IOps2<T>{static abstract T Parse(string s); static abstract T Add(T a,T b);}
struct Pair:IOps2<Pair>{public int A,B; public static Pair Parse(string s){var p=s.Split(','); return new Pair{A=int.Parse(p[0]),B=int.Parse(p[1])};} public static Pair Add(Pair a,Pair b)=>new Pair{A=a.A+b.A,B=a.B+b.B};}
var p=Pair.Parse("1,2"); __Check((Pair.Add(p,p).A).ToString(), "2");
