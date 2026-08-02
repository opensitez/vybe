// vybe-test: csharp/csharp_const_and_readonly_fields/const_field_is_accessible_without_instance
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Limits {
    public const int Max = 100;
}
__Check((Limits.Max).ToString(), "100");
