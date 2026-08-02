// vybe-test: csharp/csharp_jagged_arrays/jagged_array_element_access_uses_double_indexer
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[][] jag = new int[][]{ new[]{10,20}, new[]{30,40,50} };
__Check((jag[1][2]).ToString(), "50");
