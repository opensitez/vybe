// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_bool_try_pattern
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

interface ITry<T> where T:ITry<T>{static abstract bool Try(string s,out T value);}
struct Token:ITry<Token>{public string Raw; public static bool Try(string s,out Token value){value=new Token{Raw=s};return s.Length>0;}}
Token t; __P((Token.Try("x",out t)).ToString());
__Check("True");
