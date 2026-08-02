// vybe-test: csharp/csharp_ranges_indices/range_from_start_to_end_returns_full_array
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a={1,2,3}; var s=a[..];
__Check((s.Length).ToString(), "3");
