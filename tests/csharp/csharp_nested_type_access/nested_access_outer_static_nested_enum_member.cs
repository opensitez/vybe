// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_static_nested_enum_member
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config{public enum Level{Low,High} public static Level Default=>Level.Low;} __Check((Config.Default).ToString(), "Low");
