// vybe-test: csharp/collections_advanced/multidimensional_array
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] matrix = { { 1, 2 }, { 3, 4 }, { 5, 6 } };
__Check((matrix[0, 0]).ToString(), "1");
__Check((matrix[1, 1]).ToString(), "4");
__Check((matrix[2, 0]).ToString(), "5");
