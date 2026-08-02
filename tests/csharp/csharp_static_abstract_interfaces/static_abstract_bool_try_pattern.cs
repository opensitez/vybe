// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_bool_try_pattern
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ITry<T> where T:ITry<T>{static abstract bool Try(string s,out T value);}
struct Token:ITry<Token>{public string Raw; public static bool Try(string s,out Token value){value=new Token{Raw=s};return s.Length>0;}}
Token t; __Check((Token.Try("x",out t)).ToString(), "True");
