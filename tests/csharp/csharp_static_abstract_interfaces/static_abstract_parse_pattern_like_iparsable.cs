// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_parse_pattern_like_iparsable
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IParsable<T> where T:IParsable<T>{static abstract T Parse(string s);}
struct Age:IParsable<Age>{public int Years; public static Age Parse(string s)=>new Age{Years=int.Parse(s)};}
__Check((Age.Parse("30").Years).ToString(), "30");
