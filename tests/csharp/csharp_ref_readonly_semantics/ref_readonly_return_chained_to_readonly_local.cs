// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_return_chained_to_readonly_local
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] values={100,200}; ref readonly int Head()=>ref values[0]; ref readonly int h=ref Head(); __Check((h).ToString(), "100");
