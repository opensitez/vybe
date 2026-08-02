// vybe-test: csharp/csharp_nested_type_access/nested_access_private_nested_static_from_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Cache{static class Store{public static int V=5;} public static int Read()=>Store.V;} __Check((Cache.Read()).ToString(), "5");
