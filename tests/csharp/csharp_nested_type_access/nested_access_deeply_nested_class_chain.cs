// vybe-test: csharp/csharp_nested_type_access/nested_access_deeply_nested_class_chain
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class L1{public class L2{public class L3{public int V=42;}}} __Check((new L1.L2.L3().V).ToString(), "42");
