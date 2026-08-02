// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_parameter_accepts_array_element_reference
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Show(ref readonly int value){__Check((value).ToString(), "7");} int[] arr={7,8}; Show(ref arr[0]);
