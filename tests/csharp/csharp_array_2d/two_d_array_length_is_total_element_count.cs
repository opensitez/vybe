// vybe-test: csharp/csharp_array_2d/two_d_array_length_is_total_element_count
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] m=new int[3,4];
__Check((m.Length).ToString(), "12");
