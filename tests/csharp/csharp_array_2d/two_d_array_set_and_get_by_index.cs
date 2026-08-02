// vybe-test: csharp/csharp_array_2d/two_d_array_set_and_get_by_index
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,] m=new int[3,3];
m[1,2]=99;
__Check((m[1,2]).ToString(), "99");
