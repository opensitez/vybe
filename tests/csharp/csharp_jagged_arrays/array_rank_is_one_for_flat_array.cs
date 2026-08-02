// vybe-test: csharp/csharp_jagged_arrays/array_rank_is_one_for_flat_array
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = new int[5]; __Check((a.Rank).ToString(), "1");
