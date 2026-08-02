// vybe-test: csharp/csharp_array_advanced/multidimensional_array_length_at_each_dimension
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] grid=new int[3,4];
__Check((grid.GetLength(0)).ToString(), "3"); __Check((grid.GetLength(1)).ToString(), "4");
