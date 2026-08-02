// vybe-test: csharp/csharp_nested_type_access/nested_access_two_nested_classes_same_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pair{public class Left{public int V=1;} public class Right{public int V=2;}} __Check((new Pair.Left().V).ToString(), "1"); __Check((new Pair.Right().V).ToString(), "2");
