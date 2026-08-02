// vybe-test: csharp/csharp_ranges_indices/range_slice_returns_sub_array
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a={1,2,3,4,5}; var s=a[1..4];
__Check((s.Length).ToString(), "3"); __Check((s[0]).ToString(), "2"); __Check((s[2]).ToString(), "4");
