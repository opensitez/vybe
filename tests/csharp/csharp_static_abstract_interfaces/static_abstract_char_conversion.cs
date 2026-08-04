// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_char_conversion
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

interface IChar<T> where T:IChar<T>{static abstract T FromChar(char c);}
struct Letter:IChar<Letter>{public char C; public static Letter FromChar(char c)=>new Letter{C=c};}
__P((Letter.FromChar('z').C).ToString());
__Check("z");
