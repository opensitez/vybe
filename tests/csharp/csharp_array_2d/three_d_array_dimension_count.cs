// vybe-test: csharp/csharp_array_2d/three_d_array_dimension_count
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[,,] t=new int[2,3,4];
__Check((t.Rank).ToString(), "3");
