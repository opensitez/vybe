// vybe-test: csharp/csharp_array_2d/two_d_array_row_and_column_lengths
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] m=new int[4,5];
__Check((m.GetLength(0)).ToString(), "4"); __Check((m.GetLength(1)).ToString(), "5");
