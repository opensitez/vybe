// vybe-test: csharp/csharp_const_and_readonly_fields/static_readonly_vs_const_both_accessible_on_type_name
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Mix {
    public const int A = 1;
    public static readonly int B = 2;
}
__Check((Mix.A + Mix.B).ToString(), "3");
