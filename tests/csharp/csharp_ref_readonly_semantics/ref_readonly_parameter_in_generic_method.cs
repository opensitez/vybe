// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_parameter_in_generic_method
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static int Read<T>(ref readonly T value) where T: struct { return value.ToString().Length; } int n=123; __Check((Read(ref n)>0).ToString(), "True");
