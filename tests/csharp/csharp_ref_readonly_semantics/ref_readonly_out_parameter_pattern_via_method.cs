// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_out_parameter_pattern_via_method
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool TryGet(ref readonly int[] src,int i,out int value){value=src[i]; return true;} int[] arr={12}; TryGet(ref arr,0,out int v); __Check((v).ToString(), "12");
