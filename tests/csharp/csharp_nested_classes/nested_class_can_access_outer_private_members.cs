// vybe-test: csharp/csharp_nested_classes/nested_class_can_access_outer_private_members
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer{
    static int secret=42;
    public class Inner{public int Get()=>secret;}
}
__Check((new Outer.Inner().Get()).ToString(), "42");
