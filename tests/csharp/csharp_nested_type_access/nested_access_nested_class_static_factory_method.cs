// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_static_factory_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pool{public class Token{public int Id; public static Token Make(int id)=>new Token{Id=id};}} __Check((Pool.Token.Make(21).Id).ToString(), "21");
