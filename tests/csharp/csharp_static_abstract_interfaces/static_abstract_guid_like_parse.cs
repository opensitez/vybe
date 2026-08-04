// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_guid_like_parse
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

interface IGuid<T> where T:IGuid<T>{static abstract T Parse(string hex);}
struct Id:IGuid<Id>{public string Hex; public static Id Parse(string hex)=>new Id{Hex=hex.ToUpper()};}
__P((Id.Parse("ab").Hex).ToString());
__Check("AB");
