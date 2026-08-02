// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_conditional_selects_branch_reference
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={1,2,3}; bool pickSecond=true; ref readonly int chosen=ref (pickSecond?ref arr[1]:ref arr[0]); __Check((chosen).ToString(), "2");
