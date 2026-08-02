// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_return_from_local_function
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total=5; ref readonly int View(){return ref total;} __Check((View()).ToString(), "5");
