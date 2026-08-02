// vybe-test: csharp/csharp_static_type_behaviors/static_readonly_field_exposes_precomputed_value
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Build {
    public static readonly string Channel = "stable";
}
__Check((Build.Channel).ToString(), "stable");
