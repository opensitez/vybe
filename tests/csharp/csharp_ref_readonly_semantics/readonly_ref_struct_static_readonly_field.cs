// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_ref_struct_static_readonly_field
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly ref struct Limits{public static readonly int Max=128; public int Value;} __Check((Limits.Max).ToString(), "128");
