// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_try_parse_returns_bool
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ITryParsable<T> where T:ITryParsable<T>{static abstract bool TryParse(string s,out T value);}
struct Pair:ITryParsable<Pair>{public int A; public static bool TryParse(string s,out Pair value){value=new Pair{A=int.Parse(s)};return true;}}
Pair p; __Check((Pair.TryParse("4",out p)?p.A:-1).ToString(), "4");
