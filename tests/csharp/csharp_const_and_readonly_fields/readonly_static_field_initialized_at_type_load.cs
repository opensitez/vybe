// vybe-test: csharp/csharp_const_and_readonly_fields/readonly_static_field_initialized_at_type_load
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config {
    public static readonly string Prefix = "app";
}
__Check((Config.Prefix).ToString(), "app");
