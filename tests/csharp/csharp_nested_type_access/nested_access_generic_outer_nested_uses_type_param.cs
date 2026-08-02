// vybe-test: csharp/csharp_nested_type_access/nested_access_generic_outer_nested_uses_type_param
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T>{public class Holder{public T Value;}} var h=new Box<int>.Holder(); h.Value=15; __Check((h.Value).ToString(), "15");
