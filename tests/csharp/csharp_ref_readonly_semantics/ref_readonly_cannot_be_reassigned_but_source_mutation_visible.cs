// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_cannot_be_reassigned_but_source_mutation_visible
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={1,2}; ref readonly int r=ref arr[1]; arr[1]=50; __Check((r).ToString(), "50");
