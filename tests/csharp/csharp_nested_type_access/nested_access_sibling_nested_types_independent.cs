// vybe-test: csharp/csharp_nested_type_access/nested_access_sibling_nested_types_independent
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Duo{public class A{public int Bump(int n)=>n+1;} public class B{public int Bump(int n)=>n+2;}} __Check((new Duo.A().Bump(5)).ToString(), "6"); __Check((new Duo.B().Bump(5)).ToString(), "7");
