// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_static_method_creates_nested
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Hub{public class Node{public int V=3;} public static Node Create()=>new Node();} __Check((Hub.Create().V).ToString(), "3");
