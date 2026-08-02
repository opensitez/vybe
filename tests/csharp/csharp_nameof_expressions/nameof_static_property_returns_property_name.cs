// vybe-test: csharp/csharp_nameof_expressions/nameof_static_property_returns_property_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config{public static int Port{get;set;}=80;} __Check((nameof(Config.Port)).ToString(), "Port");
