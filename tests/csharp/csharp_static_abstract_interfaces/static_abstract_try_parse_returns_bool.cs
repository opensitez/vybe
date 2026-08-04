// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_try_parse_returns_bool
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

interface ITryParsable<T> where T:ITryParsable<T>{static abstract bool TryParse(string s,out T value);}
struct Pair:ITryParsable<Pair>{public int A; public static bool TryParse(string s,out Pair value){value=new Pair{A=int.Parse(s)};return true;}}
Pair p; __P((Pair.TryParse("4",out p)?p.A:-1).ToString());
__Check("4");
