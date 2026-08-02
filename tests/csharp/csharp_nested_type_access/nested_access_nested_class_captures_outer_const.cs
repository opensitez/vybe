// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_captures_outer_const
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer{public const string Prefix="pre"; public class Inner{public string Tag()=>Prefix+"fix";}} __Check((new Outer.Inner().Tag()).ToString(), "prefix");
