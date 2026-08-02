// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_static_class_utility
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Text{public static class Util{public static string Join(string a,string b)=>a+b;} public static string Merge()=>Util.Join("a","b");} __Check((Text.Merge()).ToString(), "ab");
