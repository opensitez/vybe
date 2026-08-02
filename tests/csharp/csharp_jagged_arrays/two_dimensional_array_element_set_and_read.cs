// vybe-test: csharp/csharp_jagged_arrays/two_dimensional_array_element_set_and_read
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] m = new int[2,2];
m[0,1] = 7;
__Check((m[0,1]).ToString(), "7");
