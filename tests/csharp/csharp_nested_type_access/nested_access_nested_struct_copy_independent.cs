// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_struct_copy_independent
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Sheet{public struct Cell{public int V;} public int Sum(){var a=new Cell(); var b=a; a.V=3; b.V=5; return a.V+b.V;}} __Check((new Sheet().Sum()).ToString(), "8");
