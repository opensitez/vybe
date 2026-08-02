// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_guid_like_parse
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IGuid<T> where T:IGuid<T>{static abstract T Parse(string hex);}
struct Id:IGuid<Id>{public string Hex; public static Id Parse(string hex)=>new Id{Hex=hex.ToUpper()};}
__Check((Id.Parse("ab").Hex).ToString(), "AB");
