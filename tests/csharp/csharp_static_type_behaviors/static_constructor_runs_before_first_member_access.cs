// vybe-test: csharp/csharp_static_type_behaviors/static_constructor_runs_before_first_member_access
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Registry {
    public static string Label;
    static Registry() { Label = "ready"; }
}
__Check((Registry.Label).ToString(), "ready");
