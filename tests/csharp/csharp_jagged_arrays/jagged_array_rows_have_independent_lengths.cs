// vybe-test: csharp/csharp_jagged_arrays/jagged_array_rows_have_independent_lengths
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[][] jag = new int[3][];
jag[0] = new int[]{1};
jag[1] = new int[]{2,3};
jag[2] = new int[]{4,5,6};
__Check((jag[2].Length).ToString(), "3");
