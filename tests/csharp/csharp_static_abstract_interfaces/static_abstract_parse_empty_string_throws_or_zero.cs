// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_parse_empty_string_throws_or_zero
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ILen<T> where T:ILen<T>{static abstract int From(string s);}
struct Size:ILen<Size>{public int N; public static int From(string s)=>s.Length;}
__Check((Size.From("abc")).ToString(), "3");
