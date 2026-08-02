// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_static_inside_nested_static
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Root{public static class A{public static class B{public static int V=13;}}} __Check((Root.A.B.V).ToString(), "13");
