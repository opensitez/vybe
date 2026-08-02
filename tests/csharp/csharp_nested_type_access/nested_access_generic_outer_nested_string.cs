// vybe-test: csharp/csharp_nested_type_access/nested_access_generic_outer_nested_string
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T>{public class Holder{public T Value;} public Holder(T v){Value=v;}} __Check((new Box<string>.Holder("ok").Value).ToString(), "ok");
