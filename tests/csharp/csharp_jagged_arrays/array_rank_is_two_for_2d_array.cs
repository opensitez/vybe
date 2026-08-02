// vybe-test: csharp/csharp_jagged_arrays/array_rank_is_two_for_2d_array
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] a = new int[2,3]; __Check((a.Rank).ToString(), "2");
