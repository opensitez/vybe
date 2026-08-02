// vybe-test: csharp/csharp_array_2d/two_d_array_initializer_syntax
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] m={{1,2,3},{4,5,6}};
__Check((m[0,2]).ToString(), "3"); __Check((m[1,0]).ToString(), "4");
