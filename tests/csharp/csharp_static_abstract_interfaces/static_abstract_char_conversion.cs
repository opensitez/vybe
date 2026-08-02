// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_char_conversion
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IChar<T> where T:IChar<T>{static abstract T FromChar(char c);}
struct Letter:IChar<Letter>{public char C; public static Letter FromChar(char c)=>new Letter{C=c};}
__Check((Letter.FromChar('z').C).ToString(), "z");
