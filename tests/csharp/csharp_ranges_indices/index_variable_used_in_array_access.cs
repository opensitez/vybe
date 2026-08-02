// vybe-test: csharp/csharp_ranges_indices/index_variable_used_in_array_access
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a={10,20,30,40,50};
System.Index i=^2;
__Check((a[i]).ToString(), "40");
