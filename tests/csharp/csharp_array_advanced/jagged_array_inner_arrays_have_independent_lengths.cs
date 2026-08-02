// vybe-test: csharp/csharp_array_advanced/jagged_array_inner_arrays_have_independent_lengths
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[][] j=new int[3][];
j[0]=new int[1]; j[1]=new int[2]; j[2]=new int[3];
__Check((j[0].Length).ToString(), "1"); __Check((j[2].Length).ToString(), "3");
