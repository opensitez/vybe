// vybe-test: csharp/csharp_readonly_members/readonly_static_field_initialized_at_class_load
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config{public static readonly string Env="prod";}
__Check((Config.Env).ToString(), "prod");
