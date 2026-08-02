// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_default_string_label
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ILabel<T> where T:ILabel<T>{static virtual string Tag=>"d"; static abstract T Make();}
struct Tag:ILabel<Tag>{public static Tag Make()=>new Tag(); public static string Tag=>"x";}
__Check((Tag.Tag).ToString(), "x");
