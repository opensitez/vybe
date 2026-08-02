// vybe-test: csharp/csharp_struct_features/default_struct_instance_has_zero_numeric_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Size { public int W, H; }
Size s = default;
__Check((s.W).ToString(), "0"); __Check((s.H).ToString(), "0");
